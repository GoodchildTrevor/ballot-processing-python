import gradio as gr
import tempfile
import os
import shutil

from ballot_processor import run_voting

ALL_NOMINATIONS = [
    "director", "actor", "actress", "actor2", "actress2",
    "original_screenplay", "adapted_screenplay", "operator", "editing",
    "soundtrack", "song", "art_direction", "costumes", "make_up",
    "effects", "sound", "stunts", "animation", "documentation", "russian",
    "live_action_short", "animated_short", "documentary_short",
    "debut", "ensemble", "using_music", "young_actor", "young_actress",
    "choreography", "special_mentions"
]

def run_processing(file_obj, selected_noms):
    if file_obj is None:
        raise gr.Error("Пожалуйста, загрузите файл!")
    if not selected_noms:
        raise gr.Error("Выберите хотя бы одну номинацию!")
    
    temp_dir = tempfile.gettempdir()
    output_path = os.path.join(temp_dir, f"result_{os.path.basename(file_obj.name)}")
    shutil.copy2(file_obj.name, output_path)

    try:
        sorted_selected = [n for n in ALL_NOMINATIONS if n in selected_noms]
        
        run_voting(output_path, sorted_selected)
        
        return output_path
    except Exception as e:
        if os.path.exists(output_path):
            os.unlink(output_path)
        raise gr.Error(f"Ошибка обработки: {str(e)}")

with gr.Blocks(title="Кинопоиск Голосование", theme=gr.themes.Soft()) as demo:
    gr.Markdown("## 🎬 Обработка бюллетеней")
    gr.Markdown(
        "Загрузите Excel-файл. Напоминание: в файле должны быть листы **'номинанты'** (голоса) и **'списки'**"
    )
    
    with gr.Row():
        with gr.Column(scale=1):
            file_input = gr.File(
                label="Загрузите Excel-файл (.xlsx)", 
                file_types=[".xlsx"]
            )
            submit_btn = gr.Button("🚀 Обработать", variant="primary")
        
        with gr.Column(scale=2):
            with gr.Row():
                select_all = gr.Button("Выбрать все", size="sm")
                deselect_all = gr.Button("Сбросить", size="sm")
            
            nominations_input = gr.CheckboxGroup(
                choices=ALL_NOMINATIONS,
                value=ALL_NOMINATIONS,
                label="Список номинаций для обработки (влияет на порядок колонок)"
            )
    
    output_file = gr.File(label="📥 Скачать результат", interactive=False)

    select_all.click(lambda: ALL_NOMINATIONS, outputs=nominations_input)
    deselect_all.click(lambda: [], outputs=nominations_input)

    submit_btn.click(
        fn=run_processing,
        inputs=[file_input, nominations_input],
        outputs=output_file
    )

if __name__ == "__main__":
    demo.launch(server_name="0.0.0.0", server_port=7860)
