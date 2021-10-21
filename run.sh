#!/bin/bash

tmux new -s f101_auto
conda activate f101_auto
voila --port=8080 --no-browser