test:
    uv run pytest
bump v:
    uv version --bump {{v}}
publish:
    git tag $(uv version --short)
    git push origin master --tags