from pathlib import Path

import pytest

from solar_forecast.jobs.lock import TrainingAlreadyRunning, exclusive_training_lock


def test_training_lock_blocks_second_model(tmp_path):
    lock = tmp_path / ".training.lock"
    with exclusive_training_lock(lock, "xgboost"):
        with pytest.raises(TrainingAlreadyRunning):
            with exclusive_training_lock(lock, "cnn_bilstm"):
                pass
    assert not lock.exists()
