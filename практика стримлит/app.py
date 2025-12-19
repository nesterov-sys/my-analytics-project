import streamlit as st
import pandas as pd
import plotly.express as px
import numpy as np
import os

@st.cache_data
def load_data():
    file_path = "data/titanic.txt"  
    if not os.path.exists(file_path):
        st.error(f"Файл {file_path} не найден. Поместите файл 'titanic.csv' в папку 'data' внутри папки проекта.")
        st.stop()
    data = pd.read_csv(file_path)
    data['Age'] = data['Age'].fillna(data['Age'].median())
    data['Survived'] = data['Survived'].map({0: "Погиб", 1: "Выжил"})
    data['Sex'] = data['Sex'].str.capitalize()
    return data

df = load_data()

st.set_page_config(page_title="Titanic Analysis", layout="wide")
st.title("Анализ пассажиров Титаника")
st.markdown("Исследование факторов выживаемости на основе публичного датасета")

st.sidebar.header("Фильтры")
selected_class = st.sidebar.multiselect(
    "Класс каюты",
    options=df['Pclass'].unique(),
    default=df['Pclass'].unique()
)
selected_sex = st.sidebar.multiselect(
    "Пол",
    options=df['Sex'].unique(),
    default=df['Sex'].unique()
)
age_range = st.sidebar.slider(
    "Возрастной диапазон",
    min_value=int(df['Age'].min()),
    max_value=int(df['Age'].max()),
    value=(int(df['Age'].min()), int(df['Age'].max()))
)

filtered_df = df[
    (df['Pclass'].isin(selected_class)) &
    (df['Sex'].isin(selected_sex)) &
    (df['Age'] >= age_range[0]) &
    (df['Age'] <= age_range[1])
]

col1, col2, col3, col4 = st.columns(4)
col1.metric("Пассажиры", len(filtered_df))
col2.metric("Выжившие", f"{len(filtered_df[filtered_df['Survived'] == 'Выжил'])} ({len(filtered_df[filtered_df['Survived'] == 'Выжил']) / len(filtered_df) * 100:.1f}%)")
col3.metric("Средний возраст", f"{filtered_df['Age'].mean():.1f}")
col4.metric("Средний тариф", f"${filtered_df['Fare'].mean():.2f}")

st.subheader("Доля выживших")
fig1 = px.pie(
    filtered_df, 
    names='Survived', 
    title='Общая выживаемость',
    hole=0.4,
    color_discrete_sequence=['#FF6B6B', '#4ECDC4']
)
st.plotly_chart(fig1, use_container_width=True)

st.subheader("Выживаемость по возрастным группам")
filtered_df['AgeGroup'] = pd.cut(filtered_df['Age'], bins=[0, 12, 18, 35, 60, 100], labels=['Дети', 'Подростки', 'Молодежь', 'Взрослые', 'Пожилые'])
survival_by_age = filtered_df.groupby('AgeGroup')['Survived'].value_counts(normalize=True).unstack().fillna(0)
fig2 = px.bar(
    survival_by_age.reset_index(),
    x='AgeGroup',
    y='Выжил',
    title='Доля выживших по возрастам',
    labels={'Выжил': 'Доля выживших'},
    color_discrete_sequence=['#4ECDC4']
)
st.plotly_chart(fig2, use_container_width=True)

st.subheader("Тариф vs Возраст")
fig3 = px.scatter(
    filtered_df,
    x='Fare',
    y='Age',
    color='Survived',
    size='Pclass',
    hover_name='Name',
    title='Зависимость выживаемости от тарифа и возраста',
    color_discrete_map={'Выжил': '#4ECDC4', 'Погиб': '#FF6B6B'}
)
st.plotly_chart(fig3, use_container_width=True)

st.subheader("Данные")
if st.checkbox("Показать исходные данные"):
    st.dataframe(filtered_df[['Name', 'Pclass', 'Sex', 'Age', 'Fare', 'Survived']].head(20))

st.download_button(
    label="Скачать данные (CSV)",
    data=filtered_df.to_csv(index=False),
    file_name="titanic_filtered.csv",
    mime="text/csv"
)
