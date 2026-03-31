import os
from itertools import combinations
import tempfile
import zipfile

import logging
import gradio as gr
import pandas as pd
import Levenshtein

from utils import define_decade, extract_year, process_results, save_df_as_image

logger = logging.getLogger("genres_app")
logger.setLevel(logging.INFO)

handler = logging.StreamHandler()
handler.setLevel(logging.INFO)
formatter = logging.Formatter(
    fmt="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
handler.setFormatter(formatter)
logger.addHandler(handler)


def process_file(excel_path: str):
    """
    Main processing entrypoint for the Gradio interface.

    :param excel_path: Path to uploaded Excel file on server
    :type excel_path: str
    :return: (zip archive path, preview image path or None, DataFrame with similar titles)
    :rtype: Tuple[str, Optional[str], pd.DataFrame]
    """
    logger.info("Received file: %s", excel_path)

    file_path = excel_path
    filename = os.path.basename(file_path)
    name = filename

    try:
        data = pd.read_excel(file_path)
        logger.info("Excel loaded, shape: %s", data.shape)
    except Exception as e:
        logger.exception("Failed to read Excel file")
        raise

    try:
        users = data['Ваш ник на Форуме Кинопоиска:']
        data = data.drop(columns=['Отметка времени', 'Ваш ник на Форуме Кинопоиска:'])
        logger.info("Columns after drop: %s", list(data.columns))
    except KeyError as e:
        logger.exception("Expected columns are missing in Excel")
        raise

    try:
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
        logger.info("Built results list: %d entries", len(results))
    except Exception:
        logger.exception("Error while building results list")
        raise

    top = pd.DataFrame(results)
    logger.info("Top DataFrame shape: %s", top.shape)

    top = top.rename(columns={
        'Фильм': 'фильмы',
        'Упоминаний': 'упоминания',
        'Баллов': 'баллы',
        'Места': 'позиция'
    })

    top = top.dropna(subset=['фильмы'])
    logger.info("Top after dropping NaN films: %s", top.shape)

    top['фильмы'] = top['фильмы'].apply(lambda x: x.strip())
    top['год'] = top['фильмы'].apply(extract_year)
    top['декада'] = top['год'].apply(define_decade)

    grouped_top = top.groupby('фильмы').agg({
        'упоминания': 'sum',
        'баллы': 'sum'
    }).reset_index()
    logger.info("Grouped top shape: %s", grouped_top.shape)

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
    logger.info("Sorted top final shape: %s", sorted_top.shape)

    # Levenshtein similar films
    similar_films = [
        (a, b) for a, b in combinations(sorted_top['фильмы'], 2)
        if Levenshtein.distance(a, b) <= 3
    ]
    logger.info("Found %d similar film pairs (Levenshtein <= 3)", len(similar_films))

    similar_films_df = pd.DataFrame(similar_films, columns=['Фильм 1', 'Фильм 2'])
    if not similar_films_df.empty:
        similar_films_df['Расстояние'] = similar_films_df.apply(
            lambda row: Levenshtein.distance(row['Фильм 1'], row['Фильм 2']), axis=1
        )
        similar_films_df = similar_films_df.sort_values('Расстояние').reset_index(drop=True)

    tmpdir = tempfile.mkdtemp()
    logger.info("Created temp dir: %s", tmpdir)

    excel_out_path = os.path.join(tmpdir, f"results_{name}.xlsx")
    sorted_top.to_excel(excel_out_path, index=False)
    logger.info("Saved Excel results to: %s", excel_out_path)

    num_lists = len(data)
    max_milestone = (num_lists // 10) * 10
    milestones = [m for m in range(10, max_milestone + 1, 10)] if max_milestone >= 10 else []
    logger.info("Num lists: %d, milestones: %s", num_lists, milestones)

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
        logger.info("Saved image for milestone %d: %s", m, img_path)

    zip_path = os.path.join(tmpdir, f"results_{name}.zip")
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        zf.write(excel_out_path, os.path.basename(excel_out_path))
        for img in image_paths:
            zf.write(img, os.path.basename(img))
    logger.info("Created zip archive: %s", zip_path)

    preview_img = image_paths[-1] if image_paths else None
    if preview_img:
        logger.info("Preview image: %s", preview_img)
    else:
        logger.info("No preview image (no milestones)")

    return zip_path, preview_img, similar_films_df


with gr.Blocks(
    title='Жанровый топ фильмов',
    theme=gr.themes.Soft()
) as demo:
    file_input = gr.File(type='filepath', label='Загрузите Excel с топами')

    submit_btn = gr.Button('Обработать')
    clear_btn = gr.Button('Очистить')

    output_zip = gr.File(label='Архив с результатами (Excel + картинки)')
    similar_df = gr.DataFrame(
        label='Похожие названия фильмов (расстояние Левенштейна ≤ 3)',
        headers=['Фильм 1', 'Фильм 2', 'Расстояние'],
        wrap=True
    )

    submit_btn.click(
        fn=process_file,
        inputs=file_input,
        outputs=[output_zip, similar_df]
    )
    clear_btn.click(lambda: None, outputs=file_input)


if __name__ == "__main__":
    logger.info("Starting Gradio genres app...")
    demo.launch(server_name="0.0.0.0", server_port=7861)
