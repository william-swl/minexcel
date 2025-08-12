test:
    uv run pytest
bump v:
    uv version --bump {{v}}

default_port := '12600'
doc_serve port=default_port:
    uv run mkdocs serve --dev-addr=0.0.0.0:{{port}}
doc_build:
    uv run mkdocs gh-deploy --clean

publish:
    uv run mkdocs gh-deploy --clean
    git tag $(uv version --short)
    git push origin master --tags
