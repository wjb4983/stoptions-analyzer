import numpy as np

from snn_bench.eval import classification_metrics
from snn_bench.feature_pipelines import BasicFeaturePipeline
from snn_bench.models import DummySpikingModel
from snn_bench.tasks import DirectionClassificationTask
from snn_bench.trainers import SimpleTrainer


def test_dummy_pipeline_shapes() -> None:
    bars = {
        "c": np.array([100.0, 101.0, 99.0, 100.0, 103.0, 105.0]),
        "t": np.arange(6),
        "o": np.zeros(6),
        "h": np.zeros(6),
        "l": np.zeros(6),
        "v": np.zeros(6),
        "n": np.zeros(6),
    }
    features = BasicFeaturePipeline().transform(bars)
    x, y = DirectionClassificationTask().make_dataset(features, bars)
    preds = SimpleTrainer(DummySpikingModel()).run(x, y, x)
    m = classification_metrics(y, preds)
    assert "accuracy" in m
