#!/bin/bash
# One-time setup on an IU Quartz login node:
#   bash quartz/setup_env.sh
# Creates a conda environment and points the Hugging Face cache at Slate
# so the large model weights do not fill your 100 GB home quota.
set -euo pipefail

SLATE=/N/slate/$USER
mkdir -p "$SLATE/hf_cache" "$SLATE/conda"

# Install Miniforge into Slate if conda is not already available.
if ! command -v conda >/dev/null 2>&1; then
  if [ ! -x "$SLATE/miniforge3/bin/conda" ]; then
    wget -q https://github.com/conda-forge/miniforge/releases/latest/download/Miniforge3-Linux-x86_64.sh -O /tmp/miniforge.sh
    bash /tmp/miniforge.sh -b -p "$SLATE/miniforge3"
  fi
  source "$SLATE/miniforge3/etc/profile.d/conda.sh"
fi

conda create -y -p "$SLATE/conda/mcard" python=3.11
conda activate "$SLATE/conda/mcard"
pip install --upgrade pip
pip install -r "$(dirname "$0")/../requirements.txt"

# Download the model weights now, on the login node (compute nodes may lack internet).
export HF_HOME=$SLATE/hf_cache
python - <<'PY'
from huggingface_hub import snapshot_download
snapshot_download("Qwen/Qwen2.5-32B-Instruct")
PY
echo "Setup complete. Environment: $SLATE/conda/mcard"
