import gradio as gr
import tempfile
import os
from processor import process_votes

NOMINATIONS = [
    "director", "actor", "actress", "actor2", "actress2",
    "original_screenplay", "adapted_screenplay", "operator", "editing",
    "soundtrack", "song", "art_direction", "costumes", "make_up",
    "effects", "sound", "stunts", "animation", "documentation", "russian",
    "live_action_short", "animated_short", "documentary_short",
    "debut", "ensemble", "using_music", "young_actor", "young_actress",
    "choreography", "special_mentions"
]

def run_processing(file, nominations):
    if not nominations:
        raise gr.Error("Выберите хотя бы одну номинацию!")
    
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_out:
        output_path = tmp_out.name

    try:
        process_votes(file.name, output_path, nominations)
        return output_path
    except Exception as e:
        if os.path.exists(output_path):
            os.unlink(output_path)
        raise gr.Error(f"Ошибка обработки: {str(e)}")

with gr.Blocks(title="Кинопоиск Голосование") as demo:
    gr.Markdown("## 🎬 Обработка бюллетеней Кинопоиска")
    gr.Markdown("Загрузите Excel-файл с листами `номинанты` и `списки`")
    
    with gr.Row():
        file_input = gr.File(label="Excel-файл (.xlsx)", file_types=[".xlsx"])
        nominations_input = gr.CheckboxGroup(
            choices=NOMINATIONS,
            value=NOMINATIONS,
            label="Номинации"
        )
    
    submit_btn = gr.Button("Обработать")
    output_file = gr.File(label="Результат", interactive=False)

    submit_btn.click(
        fn=run_processing,
        inputs=[file_input, nominations_input],
        outputs=output_file
    )

if __name__ == "__main__":
    demo.launch(server_name="0.0.0.0", server_port=7860)
