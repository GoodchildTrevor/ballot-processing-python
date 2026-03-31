import gradio as gr
import tempfile
import os
import shutil
import logging

from ballot_processor import run_voting

logger = logging.getLogger("years_app")
logger.setLevel(logging.INFO)

handler = logging.StreamHandler()
handler.setLevel(logging.INFO)
formatter = logging.Formatter(
    fmt="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
handler.setFormatter(formatter)
logger.addHandler(handler)


ALL_NOMINATIONS: list[str] = [
    "director", "actor", "actress", "actor2", "actress2",
    "original_screenplay", "adapted_screenplay", "operator", "editing",
    "soundtrack", "song", "art_direction", "costumes", "make_up",
    "effects", "sound", "stunts", "animation", "documentation", "russian",
    "live_action_short", "animated_short", "documentary_short",
    "debut", "ensemble", "using_music", "young_actor", "young_actress",
    "choreography", "special_mentions",
]


def run_processing(file_obj, selected_noms: list[str]) -> str:
    """
    Process uploaded Excel file with voting data and return path to result file.

    :param file_obj: Uploaded file object from Gradio (File component)
    :type file_obj: gradio.File or None
    :param selected_noms: List of selected nomination keys
    :type selected_noms: List[str]
    :return: Path to processed Excel file
    :rtype: str
    """
    logger.info("run_processing called")

    if file_obj is None:
        logger.warning("No file uploaded")
        raise gr.Error("Пожалуйста, загрузите файл!")
    if not selected_noms:
        logger.warning("No nominations selected")
        raise gr.Error("Выберите хотя бы одну номинацию!")

    logger.info("Uploaded file: %s", getattr(file_obj, "name", "<?>"))
    logger.info("Selected nominations: %s", selected_noms)

    temp_dir: str = tempfile.gettempdir()
    output_path: str = os.path.join(
        temp_dir, f"result_{os.path.basename(file_obj.name)}"
    )
    logger.info("Temp dir: %s, output file: %s", temp_dir, output_path)

    shutil.copy2(file_obj.name, output_path)
    logger.info("Copied uploaded file to temp location")

    try:
        sorted_selected: list[str] = [
            n for n in ALL_NOMINATIONS if n in selected_noms
        ]
        logger.info("Sorted nominations (preserving canonical order): %s", sorted_selected)

        run_voting(output_path, sorted_selected)
        logger.info("run_voting finished successfully")

        return output_path
    except Exception as e:
        logger.exception("Error during run_voting")
        if os.path.exists(output_path):
            try:
                os.unlink(output_path)
                logger.info("Removed temporary file after error: %s", output_path)
            except Exception:
                logger.exception("Failed to remove temporary file: %s", output_path)
        raise gr.Error(f"Ошибка обработки: {str(e)}")


with gr.Blocks(title="Кинопоиск Голосование", theme=gr.themes.Soft()) as demo:
    gr.Markdown("## 🎬 Обработка бюллетеней")
    gr.Markdown(
        "Загрузите Excel-файл. Напоминание: в файле должны быть листы "
        "**'номинанты'** (голоса) и **'списки'**"
    )

    with gr.Row():
        with gr.Column(scale=1):
            file_input = gr.File(
                label="Загрузите Excel-файл (.xlsx)",
                file_types=[".xlsx"],
            )
            submit_btn = gr.Button("🚀 Обработать", variant="primary")

        with gr.Column(scale=2):
            with gr.Row():
                select_all = gr.Button("Выбрать все", size="sm")
                deselect_all = gr.Button("Сбросить", size="sm")

            nominations_input = gr.CheckboxGroup(
                choices=ALL_NOMINATIONS,
                value=ALL_NOMINATIONS,
                label=(
                    "Список номинаций для обработки "
                    "(влияет на порядок колонок)"
                ),
            )

    output_file = gr.File(label="📥 Скачать результат", interactive=False)

    select_all.click(lambda: ALL_NOMINATIONS, outputs=nominations_input)
    deselect_all.click(lambda: [], outputs=nominations_input)

    submit_btn.click(
        fn=run_processing,
        inputs=[file_input, nominations_input],
        outputs=output_file,
    )


if __name__ == "__main__":
    logger.info("Starting Gradio years app...")
    demo.launch(server_name="0.0.0.0", server_port=7860)
