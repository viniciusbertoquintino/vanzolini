from pptx import Presentation

def generate_ppt(template_path, content_path, output_path):
    prs = Presentation(template_path)

    with open(content_path, "r", encoding="utf-8") as f:
        raw = f.read()

    sections = [s.strip() for s in raw.split("#") if s.strip()]

    for slide, section in zip(prs.slides, sections):
        lines = [l.strip() for l in section.splitlines() if l.strip()]

        if len(lines) < 2:
            continue

        titulo = lines[0]
        conteudo = "\n".join(lines[1:])

        # TÍTULO
        if slide.shapes.title:
            slide.shapes.title.text = titulo

        # CORPO (placeholder padrão: layout "Título e Conteúdo")
        try:
            slide.placeholders[1].text = conteudo
        except KeyError:
            pass

    prs.save(output_path)
