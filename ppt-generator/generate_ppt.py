from pptx import Presentation
import sys

template_path = sys.argv[1]
content_path = sys.argv[2]
output_path = sys.argv[3]

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

    # CORPO (placeholder padrão)
    try:
        body = slide.placeholders[1]
        body.text = conteudo
    except KeyError:
        pass  # slide sem corpo

prs.save(output_path)
