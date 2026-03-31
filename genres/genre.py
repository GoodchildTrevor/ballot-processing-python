import os
import zipfile
import tempfile
import gradio as gr
import pandas as pd
import Levenshtein
from itertools import combinations
from utils import define_decade, extract_year, process_results, save_df_as_image


def process_file(excel_path: str):
    file_path = excel_path
    filename = os.path.basename(file_path)
    name = filename

    data = pd.read_excel(file_path)
    users = data['Ваш ник на Форуме Кинопоиска:']
    data = data.drop(columns=['Отметка времени', 'Ваш ник на Форуме Кинопоиска:'])

    results = [
        {
            'user': idx,
            'пользователь': users[idx],
            'Фильм': row[f"Лучший фильм {col_num+1} место"],
            'Упоминаний': 1,
            'Баллов': 25 - col_num,
            'Места': col_num + 1
        }
        for idx, row in data.iterrows()
        for col_num in range(25)
    ]

    top = pd.DataFrame(results)
    top = top.rename(columns={
        'Фильм': 'фильмы',
        'Упоминаний': 'упоминания',
        'Баллов': 'баллы',
        'Места': 'позиция'
    })

    top = top.dropna(subset=['фильмы'])
    top['фильмы'] = top['фильмы'].apply(lambda x: x.strip())
    top['год'] = top['фильмы'].apply(extract_year)
    top['декада'] = top['год'].apply(define_decade)

    grouped_top = top.groupby('фильмы').agg({
        'упоминания': 'sum',
        'баллы': 'sum'
    }).reset_index()

    grouped_positions = top.groupby('фильмы')['позиция'].apply(list).reset_index()

    merged_group = pd.merge(grouped_top, grouped_positions, on='фильмы')
    merged_group['позиции'] = merged_group['позиция'].apply(lambda x: tuple(sorted(x)))

    sorted_top = merged_group.sort_values(
        by=['баллы', 'упоминания', 'позиции'],
        ascending=[False, False, True]
    )
    sorted_top = sorted_top.drop(columns=['позиция']).reset_index(drop=True)
    sorted_top = pd.merge(
        sorted_top,
        top[['фильмы', 'декада']],
        on='фильмы',
        how='left'
    ).drop_duplicates()

    similar_films = [
        (a, b) for a, b in combinations(sorted_top['фильмы'], 2)
        if Levenshtein.distance(a, b) <= 3
    ]
    similar_films_df = pd.DataFrame(similar_films, columns=['Фильм 1', 'Фильм 2'])
    similar_films_df['Расстояние'] = similar_films_df.apply(
        lambda row: Levenshtein.distance(row['Фильм 1'], row['Фильм 2']), axis=1
    )
    similar_films_df = similar_films_df.sort_values('Расстояние').reset_index(drop=True)

    tmpdir = tempfile.mkdtemp()

    excel_out_path = os.path.join(tmpdir, f"results_{name}.xlsx")
    sorted_top.to_excel(excel_out_path, index=False)

    num_lists = len(data)
    max_milestone = (num_lists // 10) * 10
    milestones = [m for m in range(10, max_milestone + 1, 10)] if max_milestone >= 10 else []

    image_paths = []
    for m in milestones:
        subset_of_results = results[:m * 25]
        interim_results = process_results(subset_of_results)
        interim_results = interim_results.reset_index(drop=True)
        interim_results.index = interim_results.index + 1

        img_path = os.path.join(tmpdir, f"results_after_{m}.png")
        save_df_as_image(
            interim_results.head(10).drop(columns='Позиции в списках'),
            img_path
        )
        image_paths.append(img_path)

    zip_path = os.path.join(tmpdir, f"results_{name}.zip")
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        zf.write(excel_out_path, os.path.basename(excel_out_path))
        for img in image_paths:
            zf.write(img, os.path.basename(img))

    preview_img = image_paths[-1] if image_paths else None

    return zip_path, preview_img, similar_films_df

demo = gr.Interface(
    fn=process_file,
    inputs=gr.File(type='filepath', label='Загрузите Excel с топами'),
    outputs=[
        gr.File(label='Архив с результатами (Excel + картинки)'),
        gr.Image(label='Превью последнего майлстоуна', type='filepath'),
        gr.DataFrame(
            label='Похожие названия фильмов (расстояние Левенштейна ≤ 3)',
            headers=['Фильм 1', 'Фильм 2', 'Расстояние'],
            wrap=True
        )
    ],
    title='Жанровый топ фильмов',
    description='Загружаете Excel — получаете общий топ и промежуточные результаты в одном архиве'
)

if __name__ == "__main__":
    demo.launch(server_name="0.0.0.0", server_port=7861)
