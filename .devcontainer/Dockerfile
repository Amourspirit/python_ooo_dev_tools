FROM dockdock/libre_office_script_provider:debian-11.7-slim-1 as compiler

# ARG PYTHON_VERSION=3.9

ENV PYTHONUNBUFFERED 1

WORKDIR /workspace/ooouno-dev-tools

COPY ./pyproject.toml ./poetry.lock ./poetry.toml ./

RUN poetry install --no-root
# /workspace/ooouno-dev-tools

FROM dockdock/libre_office_script_provider:debian-11.7-slim-1 as runner

WORKDIR /workspace/ooouno-dev-tools
COPY --from=compiler /workspace/ooouno-dev-tools/.venv ./.venv

ENV VIRTUAL_ENV=/workspace/ooouno-dev-tools/.venv
ENV PATH="$VIRTUAL_ENV/bin:$PATH"

# RUN oooenv cmd-link -a

WORKDIR /home/user

COPY ./tests/fixtures/ ./

ARG WD=/workspace/ooouno-dev-tools/.venv/lib/python3.9/site-packages
WORKDIR ${WD}

RUN ln -s /usr/lib/python3/dist-packages/uno.py uno.py \
    && ln -s /usr/lib/python3/dist-packages/unohelper.py unohelper.py \
    && echo "/workspace/ooouno-dev-tools" >> ooo_dev_tools.pth
# Enable venv
# ENV PATH="/opt/venv/bin:$PATH"