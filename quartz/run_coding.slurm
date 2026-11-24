#!/bin/bash
# Submit from the repository root on Quartz:
#   sbatch quartz/run_coding.slurm
#
# Edit the --account line to your Slurm allocation (find it with: sacctmgr show assoc user=$USER).
#SBATCH --job-name=mcard-coding
#SBATCH --account=REPLACE_WITH_YOUR_ALLOCATION
#SBATCH --partition=gpu
#SBATCH --nodes=1
#SBATCH --gpus-per-node=4
#SBATCH --cpus-per-task=16
#SBATCH --mem=200G
#SBATCH --time=12:00:00
#SBATCH --output=output/slurm-%j.out
#SBATCH --mail-type=END,FAIL

set -euo pipefail
SLATE=/N/slate/$USER
[ -f "$SLATE/miniforge3/etc/profile.d/conda.sh" ] && source "$SLATE/miniforge3/etc/profile.d/conda.sh"
conda activate "$SLATE/conda/mcard"

export HF_HOME=$SLATE/hf_cache
export HF_HUB_OFFLINE=1          # weights were downloaded by setup_env.sh
export VLLM_WORKER_MULTIPROC_METHOD=spawn

# Use the first codebook file found in data/codebook/
CODEBOOK=$(find data/codebook -type f ! -name '.gitkeep' | head -n 1)
STAMP=$(date +%Y%m%d_%H%M)

nvidia-smi
python code_model_cards.py \
  --cards data/model_cards \
  --codebook "$CODEBOOK" \
  --out "output/model_card_coding_${STAMP}.xlsx" \
  --backend vllm \
  --model Qwen/Qwen2.5-32B-Instruct \
  --tp 4 \
  --dtype float16
