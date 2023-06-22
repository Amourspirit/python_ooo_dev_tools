FROM dockdock/libre_office_script_provider:22.04-UBUNTU as compiler
# FROM mcr.microsoft.com/devcontainers/python:0-3.11

ARG PYTHON_VERSION="3.10"

ENV PYTHONUNBUFFERED 1

WORKDIR /app

COPY ./pyproject.toml ./poetry.lock ./poetry.toml ./

RUN poetry install --no-root
# /workspace/ooouno-dev-tools

FROM dockdock/libre_office_script_provider:22.04-UBUNTU as runner

WORKDIR /workspace/ooouno-dev-tools
COPY --from=compiler /app/.venv ./.venv

ENV VIRTUAL_ENV=/workspace/ooouno-dev-tools/.venv
ENV PATH="$VIRTUAL_ENV/bin:$PATH"

# RUN oooenv cmd-link -a

WORKDIR /workspace/ooouno-dev-tools/.venv/lib/python$(PYTHON_VERSION)/site-packages

RUN ln -s /usr/lib/python3/dist-packages/uno.py uno.py \
    && ln -s /usr/lib/python3/dist-packages/unohelper.py unohelper.py
# Enable venv
# ENV PATH="/opt/venv/bin:$PATH"