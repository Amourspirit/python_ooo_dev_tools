FROM dockdock/libreoffice-api-dev:LO-7.5 as builder

USER root

WORKDIR /workspace/ooouno-dev-tools

COPY ./pyproject.toml ./poetry.lock ./poetry.toml ./

RUN poetry install --no-root


FROM dockdock/libreoffice-api-dev:LO-7.5 as runner

ARG PY_VER=3.10
ARG LOCAL_GIT_AUTHOR_NAME=""
ARG LOCAL_GIT_AUTHOR_EMAIL=""

USER root
ENV HOME /root

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
ENV PYTHONUNBUFFERED 1

RUN if [ x"$LOCAL_GIT_AUTHOR_NAME" != "x" ]; then git config --global user.name "$LOCAL_GIT_AUTHOR_NAME" || true; fi \
    && if [ x"$LOCAL_GIT_AUTHOR_EMAIL" != "x" ]; then git config --global user.email $LOCAL_GIT_AUTHOR_EMAIL || true; fi

WORKDIR $HOME

COPY ./.devcontainer/custom_bash.txt ./.bashrc_custom

ARG WD=/workspace/ooouno-dev-tools/.venv/lib/python$PY_VER/site-packages

WORKDIR ${WD}

RUN ln -s /usr/lib/python3/dist-packages/uno.py uno.py \
    && ln -s /usr/lib/python3/dist-packages/unohelper.py unohelper.py \
    && echo "/workspace/ooouno-dev-tools" >> ooo_dev_tools.pth

COPY ./.devcontainer/res $HOME/.tmp/res
COPY ./.devcontainer/startup.sh /usr/local/bin/startup.sh

RUN chmod -R +x "$HOME/.tmp/res/scripts" \
    && mv "$HOME/.tmp/res/scripts/unlock.sh" "/usr/local/bin/unlock" \
    && chmod +x "/usr/local/bin/unlock" \
    && mkdir -p "$HOME/.config/libreoffice/4/user/" \
    && mv "$HOME/.tmp/res/defaults/registrymodifications.xcu" "$HOME/.config/libreoffice/4/user/registrymodifications.xcu" \
    && mkdir -p "/usr/lib/libreoffice/share/extensions" \
    && chmod +x "/usr/local/bin/startup.sh"

