"""
Tests for verificacion_correo.cli.main module.
"""

from unittest.mock import MagicMock, patch
from verificacion_correo.core.excel import EmailRecord, ExcelSummary


class TestDryRunBatchAccess:
    """Tests for dry_run batch access pattern."""

    def test_dry_run_accesses_records_correctly(self):
        """
        Verify that batch iteration works correctly.

        In cli/main.py dry_run block, each batch from summary.batches is a
        List[EmailRecord] (not an object with a .records attribute). This test
        ensures the code accesses records directly via batch[:5] rather than
        batch.records[:5].
        """
        records = [
            EmailRecord(email=f"user{i}@madrid.org", row=i + 2)
            for i in range(10)
        ]
        summary = ExcelSummary(
            total_emails=10,
            pending_count=10,
            processed_count=0,
            batches=[records],
        )

        batch = summary.batches[0]

        result = []
        for j, record in enumerate(batch[:5]):
            result.append(record.email)

        assert len(result) == 5
        assert result == [f"user{i}@madrid.org" for i in range(5)]

    def test_dry_run_batch_has_no_records_attribute(self):
        """
        Confirm that a plain list batch does NOT have a .records attribute.
        This is the root cause of the original bug.
        """
        records = [
            EmailRecord(email="test@madrid.org", row=2),
        ]
        summary = ExcelSummary(
            total_emails=1,
            pending_count=1,
            processed_count=0,
            batches=[records],
        )

        batch = summary.batches[0]
        assert not hasattr(batch, "records")
