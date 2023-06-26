FROM dockdock/libreoffice-api-dev:LO-7.5 as builder

USER root

WORKDIR /workspace/ooouno-dev-tools

COPY ./pyproject.toml ./poetry.lock ./poetry.toml ./

RUN poetry install --no-root


FROM dockdock/libreoffice-api-dev:LO-7.5 as runner

ARG PY_VER=3.10
ARG GIT_AUTHOR_NAME=""
ARG GIT_AUTHOR_EMAIL=""

USER root

RUN apt-get update \
    && apt-get install -y --no-install-recommends \
    nano \
    git \
    gnupg \
    jq \
    make \
    openssh-client \
    unzip \
    wget \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*


WORKDIR /workspace/ooouno-dev-tools

# ENV PYTHONUNBUFFERED 1

COPY --from=builder /workspace/ooouno-dev-tools/.venv ./.venv

ENV VIRTUAL_ENV=/workspace/ooouno-dev-tools/.venv
ENV PATH="$VIRTUAL_ENV/bin:$PATH"

RUN git config --global user.name "$GIT_AUTHOR_NAME" \
    && git config --global user.email $GIT_AUTHOR_EMAIL

ARG WD=/workspace/ooouno-dev-tools/.venv/lib/python$PY_VER/site-packages

WORKDIR ${WD}

RUN ln -s /usr/lib/python3/dist-packages/uno.py uno.py \
    && ln -s /usr/lib/python3/dist-packages/unohelper.py unohelper.py \
    && echo "/workspace/ooouno-dev-tools" >> ooo_dev_tools.pth
# Enable venv
# ENV PATH="/opt/venv/bin:$PATH"

