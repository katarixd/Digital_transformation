import pandas as pd
import requests


# Функция для загрузки и проверки данных
def load_and_check_data(criteria_path, domains_path, values_path):
    criteria_df = pd.read_excel(criteria_path)
    domains_df = pd.read_excel(domains_path)
    values_df = pd.read_excel(values_path)

    # Объединение данных по критериям и доменам
    criteria_df = pd.merge(criteria_df, values_df, on='Критерий', how='left')
    criteria_df = pd.merge(criteria_df, domains_df, on='Домен', how='left')

    # Проверка на пропуски данных
    if criteria_df[['Вес критерия (%)', 'Вес домена (%)', 'Значение критерия (%)']].isnull().any().any():
        raise ValueError("Некоторые данные отсутствуют. Проверьте правильность всех значений в Excel-файлах.")

    # Перевод процентов в доли
    criteria_df['Вес критерия'] = criteria_df['Вес критерия (%)'] / 100
    criteria_df['Вес домена'] = criteria_df['Вес домена (%)'] / 100
    criteria_df['Значение критерия'] = criteria_df['Значение критерия (%)'] / 100

    return criteria_df


# Функция для расчета итогового уровня цифровизации
def calculate_final_score(criteria_df):
    criteria_df['Вклад критерия'] = criteria_df['Вес критерия'] * criteria_df['Вес домена'] * criteria_df[
        'Значение критерия']
    final_score = criteria_df['Вклад критерия'].sum() * 100
    return final_score


# Функция для определения интерпретации уровня цифровизации
def interpret_score(final_score):
    if final_score <= 30:
        return "Очень низкий уровень цифровизации"
    elif final_score <= 50:
        return "Низкий уровень цифровизации"
    elif final_score <= 80:
        return "Средний уровень цифровизации"
    else:
        return "Высокий уровень цифровизации"


# Формирование строки для вывода критериев
def format_criteria_data(criteria_df):
    return "\n".join([f"{row['Критерий']}: {row['Значение критерия (%)']}%" for _, row in criteria_df.iterrows()])


# Функция для форматирования итогового вывода
def format_final_output(final_score, level, criteria_df):
    output = f"""
Итоговый уровень цифровизации: {final_score:.2f}%
Интерпретация: {level}

=== Критерии и их значения ===
"""

    # Добавляем информацию о каждом критерии
    for _, row in criteria_df.iterrows():
        output += f"- **{row['Критерий']}**: {row['Значение критерия (%)']}%\n"

    return output.strip()


# Функция для анализа с использованием AI
def get_ai_advice(data, score, level):
    binary_keywords = [
        "наличие", "есть ли", "осуществляется ли", "доступность", "присутствует", "реализована", "поддерживается",
        "имеется", "отвечает ли", "предусмотрена ли"
    ]

    def is_binary(criterion):
        return any(kw in criterion.lower() for kw in binary_keywords)

    def describe_value(v):
        if v >= 90:
            return "высокий уровень"
        elif v >= 70:
            return "достаточный уровень"
        elif v >= 50:
            return "средний уровень"
        elif v >= 30:
            return "низкий уровень"
        else:
            return "очень низкий уровень"

    # Формируем данные для запроса
    data_str = "\n".join([
        f"- {row['Критерий']} — {row['Значение критерия (%)']}% ({'бинарный критерий' if is_binary(row['Критерий']) else describe_value(row['Значение критерия (%)'])})"
        for _, row in data.iterrows()
    ])

    prompt = f"""
    Ты — эксперт по цифровой трансформации. Проанализируй уровень цифровизации компании на основе следующих критериев и их значений в процентах.

    ВАЖНО:
    - Не придумывай новых критериев. Используй **только те**, что перечислены ниже.
    - Если все значения высокие и слабых сторон **нет**, напиши об этом прямо.
    - Не считай критерии со значениями 80–90% «слабыми» в общем контексте.
    - Бинарные критерии (со значениями 0% или 100%) **не считаются автоматически слабыми или сильными**. Анализируй их с учётом контекста.
    - Сравни все критерии между собой и выдели:
      - 1–3 **относительно сильных** направлений.
      - 1–2 направления, **в которых ещё есть потенциал для усиления** (не обязательно «слабые»).

    === ДАННЫЕ ===
    Общий уровень цифровизации: {score:.2f}% ({level})

    Критерии:
    {data_str}

    === ОТВЕТ ===

    СИЛЬНЫЕ СТОРОНЫ:
    Определи, какие направления компании на данный момент можно считать наиболее сильными с точки зрения цифровизации:
    - Критерий (XX%) — почему он сильный
    ...

    НАПРАВЛЕНИЯ С ПОТЕНЦИАЛОМ ДЛЯ УСИЛЕНИЯ:
    Определи, какие направления компании могут быть улучшены. Объясни, что стоит развивать дальше:
    - Критерий (XX%) — почему стоит уделить дополнительное внимание (не обязательно слабый)
    ...

    РЕКОМЕНДАЦИИ:
    1. ...
    2. ...
    3. ...
    """.strip()

    try:
        response = requests.post(
            "http://localhost:11434/api/generate",
            json={"model": "mistral", "prompt": prompt, "stream": False, "options": {"temperature": 0.7}}
        )
        return response.json().get("response", "Ошибка: Невозможно получить ответ.")
    except Exception as e:
        return f"Ошибка: {e}"


# Основной блок
def main():
    criteria_path = 'criteria.xlsx'
    domains_path = 'domains.xlsx'
    values_path = 'values_template.xlsx'

    criteria_df = load_and_check_data(criteria_path, domains_path, values_path)

    final_score = calculate_final_score(criteria_df)
    level = interpret_score(final_score)

    formatted_output = format_final_output(final_score, level, criteria_df)
    print(formatted_output)

    ai_advice = get_ai_advice(
        criteria_df[['Домен', 'Критерий', 'Значение критерия (%)']],
        final_score,
        level
    )
    print("\nАнализ Mistral-7B:")
    print(ai_advice)


if __name__ == "__main__":
    main()
