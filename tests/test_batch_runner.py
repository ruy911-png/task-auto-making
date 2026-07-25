from batch.runner import BatchState, ConditionStatus, retry_failed, run_batch


def test_run_batch_continues_after_failure(tmp_path):
    conditions = [{"x": 1}, {"x": 2}, {"x": 3}]

    def execute(condition):
        if condition["x"] == 2:
            raise RuntimeError("boom")
        return {"value": condition["x"] * 10}

    state_path = tmp_path / "job_state.json"
    state = run_batch("job1", conditions, execute, state_path)

    assert state.summary() == {"total": 3, "success": 2, "failed": 1}
    failed = state.failed_items[0]
    assert failed.condition == {"x": 2}
    assert failed.error == "boom"


def test_retry_failed_only_reruns_failed(tmp_path):
    conditions = [{"x": 1}, {"x": 2}]
    attempts = {"count": 0}

    def flaky_execute(condition):
        if condition["x"] == 2:
            attempts["count"] += 1
            if attempts["count"] == 1:
                raise RuntimeError("first try fails")
        return {"ok": True}

    state_path = tmp_path / "job_state.json"
    run_batch("job1", conditions, flaky_execute, state_path)

    state = retry_failed(flaky_execute, state_path)

    assert state.summary() == {"total": 2, "success": 2, "failed": 0}
    assert attempts["count"] == 2  # x=1은 재실행되지 않고, x=2만 한 번 더 실행됨


def test_state_roundtrip(tmp_path):
    state_path = tmp_path / "job_state.json"
    run_batch("job1", [{"x": 1}], lambda c: {"ok": True}, state_path)

    loaded = BatchState.load(state_path)

    assert loaded.job_id == "job1"
    assert loaded.items[0].status == ConditionStatus.SUCCESS
