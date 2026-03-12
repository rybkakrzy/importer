"""
API Tests for Document Storage with Versioning
Tests cover: upload, save version, restore version, get document, get versions
"""
import pytest
import requests
import base64
from uuid import uuid4


class TestDocumentStorageAPI:
    """Tests for Document Storage API endpoints"""

    @pytest.fixture(autouse=True)
    def setup(self, api_base_url):
        """Setup test data"""
        self.base_url = f"{api_base_url}/api/documentstorage"
        self.test_content = b"Hello World - Test Document"
        self.test_content_base64 = base64.b64encode(self.test_content).decode('utf-8')

    def test_upload_document_success(self):
        """Test: Upload new document returns master and version IDs"""
        payload = {
            "name": "test_upload.txt",
            "mimeType": "text/plain",
            "content": self.test_content_base64,
            "createdBy": "E2ETestUser"
        }

        response = requests.post(f"{self.base_url}/upload", json=payload)

        assert response.status_code == 200
        result = response.json()
        assert "masterId" in result
        assert "versionId" in result
        assert result["fileName"] == "test_upload.txt"
        assert result["masterId"] != result["versionId"]

        # Validate GUID format
        import uuid
        uuid.UUID(result["masterId"])
        uuid.UUID(result["versionId"])

    def test_upload_document_invalid_mime_type(self):
        """Test: Upload with invalid MIME type returns 400"""
        payload = {
            "name": "invalid.txt",
            "mimeType": "invalid-mime",  # Nieprawidłowy format
            "content": self.test_content_base64,
            "createdBy": "Tester"
        }

        response = requests.post(f"{self.base_url}/upload", json=payload)

        assert response.status_code == 400
        assert "MIME" in response.text or "nieprawidłowy" in response.text.lower()

    def test_upload_and_retrieve_document(self):
        """Test: Upload document and retrieve active version"""
        # Upload
        upload_payload = {
            "name": "retrieve_test.pdf",
            "mimeType": "application/pdf",
            "content": self.test_content_base64,
            "createdBy": "APITest"
        }
        upload_response = requests.post(f"{self.base_url}/upload", json=upload_payload)
        assert upload_response.status_code == 200
        master_id = upload_response.json()["masterId"]

        # Retrieve
        get_response = requests.get(f"{self.base_url}/{master_id}")
        
        assert get_response.status_code == 200
        document = get_response.json()
        assert document["masterId"] == master_id
        assert document["name"] == "retrieve_test.pdf"
        assert document["versionNumber"] == 1
        assert document["content"] == self.test_content_base64

    def test_save_new_version(self):
        """Test: Save new version increments version number"""
        # Upload initial document
        upload_payload = {
            "name": "versioning_test.docx",
            "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            "content": base64.b64encode(b"Version 1").decode(),
            "createdBy": "Writer"
        }
        upload_response = requests.post(f"{self.base_url}/upload", json=upload_payload)
        master_id = upload_response.json()["masterId"]

        # Save version 2
        v2_content = base64.b64encode(b"Version 2 - Updated").decode()
        save_payload = {
            "content": v2_content,
            "createdBy": "Editor"
        }
        save_response = requests.post(f"{self.base_url}/{master_id}/save", json=save_payload)

        assert save_response.status_code == 200
        result = save_response.json()
        assert result["versionNumber"] == 2
        assert "versionId" in result

        # Verify active version is v2
        get_response = requests.get(f"{self.base_url}/{master_id}")
        document = get_response.json()
        assert document["versionNumber"] == 2
        assert document["content"] == v2_content

    def test_get_document_versions_history(self):
        """Test: Get all versions returns sorted list with metadata"""
        # Upload initial
        upload_payload = {
            "name": "history_test.txt",
            "mimeType": "text/plain",
            "content": base64.b64encode(b"V1").decode(),
            "createdBy": "User"
        }
        upload_response = requests.post(f"{self.base_url}/upload", json=upload_payload)
        master_id = upload_response.json()["masterId"]

        # Add 2 more versions
        for i in range(2, 4):
            save_payload = {
                "content": base64.b64encode(f"V{i}".encode()).decode(),
                "createdBy": f"User{i}"
            }
            requests.post(f"{self.base_url}/{master_id}/save", json=save_payload)

        # Get versions
        versions_response = requests.get(f"{self.base_url}/{master_id}/versions")

        assert versions_response.status_code == 200
        versions = versions_response.json()
        assert len(versions) == 3
        
        # Sprawdź sortowanie (od najnowszej)
        assert versions[0]["versionNumber"] == 3
        assert versions[1]["versionNumber"] == 2
        assert versions[2]["versionNumber"] == 1

        # Tylko wersja 3 powinna być aktywna
        assert versions[0]["isActive"] is True
        assert versions[1]["isActive"] is False
        assert versions[2]["isActive"] is False

        # Sprawdź że nie ma contentu (tylko metadane)
        for version in versions:
            assert "content" not in version
            assert "sizeInBytes" in version
            assert version["sizeInBytes"] > 0

    def test_restore_previous_version(self):
        """Test: Restore old version makes it active again"""
        # Upload and create 3 versions
        upload_payload = {
            "name": "restore_test.md",
            "mimeType": "text/markdown",
            "content": base64.b64encode(b"Original draft").decode(),
            "createdBy": "Author"
        }
        upload_response = requests.post(f"{self.base_url}/upload", json=upload_payload)
        master_id = upload_response.json()["masterId"]
        v1_id = upload_response.json()["versionId"]

        # Version 2
        requests.post(f"{self.base_url}/{master_id}/save", json={
            "content": base64.b64encode(b"Bad edit").decode(),
            "createdBy": "Author"
        })

        # Version 3
        requests.post(f"{self.base_url}/{master_id}/save", json={
            "content": base64.b64encode(b"Worse edit").decode(),
            "createdBy": "Author"
        })

        # Restore to version 1
        restore_response = requests.post(f"{self.base_url}/{master_id}/restore/{v1_id}")

        assert restore_response.status_code == 200
        result = restore_response.json()
        assert "message" in result
        assert result["versionId"] == v1_id

        # Verify version 1 is now active
        get_response = requests.get(f"{self.base_url}/{master_id}")
        document = get_response.json()
        assert document["activeVersionId"] == v1_id
        assert document["content"] == base64.b64encode(b"Original draft").decode()

        # Verify all 3 versions still exist
        versions_response = requests.get(f"{self.base_url}/{master_id}/versions")
        versions = versions_response.json()
        assert len(versions) == 3

    def test_restore_nonexistent_version(self):
        """Test: Restore non-existent version returns 400"""
        # Upload document
        upload_payload = {
            "name": "test.txt",
            "mimeType": "text/plain",
            "content": self.test_content_base64,
            "createdBy": "User"
        }
        upload_response = requests.post(f"{self.base_url}/upload", json=upload_payload)
        master_id = upload_response.json()["masterId"]

        # Try to restore random UUID
        fake_version_id = str(uuid4())
        restore_response = requests.post(f"{self.base_url}/{master_id}/restore/{fake_version_id}")

        assert restore_response.status_code == 400
        assert "nie należy do dokumentu" in restore_response.text or "not found" in restore_response.text.lower()

    def test_get_nonexistent_document(self):
        """Test: Get non-existent document returns 404"""
        fake_master_id = str(uuid4())
        response = requests.get(f"{self.base_url}/{fake_master_id}")

        assert response.status_code == 404

    def test_save_version_to_nonexistent_document(self):
        """Test: Save version to non-existent document returns 400"""
        fake_master_id = str(uuid4())
        payload = {
            "content": self.test_content_base64,
            "createdBy": "User"
        }
        response = requests.post(f"{self.base_url}/{fake_master_id}/save", json=payload)

        assert response.status_code == 400

    def test_upload_empty_content_validation(self):
        """Test: Upload with empty content should fail validation"""
        payload = {
            "name": "empty.txt",
            "mimeType": "text/plain",
            "content": "",  # Pusty content
            "createdBy": "User"
        }
        response = requests.post(f"{self.base_url}/upload", json=payload)

        # Walidator powinien odrzucić
        assert response.status_code == 400

    def test_version_metadata_accuracy(self):
        """Test: Version metadata (size, timestamps) are accurate"""
        content1 = b"Short"
        content2 = b"This is a much longer content for testing size calculation"

        upload_payload = {
            "name": "metadata_test.bin",
            "mimeType": "application/octet-stream",
            "content": base64.b64encode(content1).decode(),
            "createdBy": "Tester"
        }
        upload_response = requests.post(f"{self.base_url}/upload", json=upload_payload)
        master_id = upload_response.json()["masterId"]

        # Add second version
        save_payload = {
            "content": base64.b64encode(content2).decode(),
            "createdBy": "Tester"
        }
        requests.post(f"{self.base_url}/{master_id}/save", json=save_payload)

        # Get versions
        versions_response = requests.get(f"{self.base_url}/{master_id}/versions")
        versions = versions_response.json()

        # Sprawdź rozmiary
        assert versions[0]["sizeInBytes"] == len(content2)  # Najnowsza
        assert versions[1]["sizeInBytes"] == len(content1)  # Pierwsza

        # Sprawdź CreatedBy
        assert all(v["createdBy"] == "Tester" for v in versions)


@pytest.mark.performance
class TestDocumentStoragePerformance:
    """Performance tests for document storage"""

    def test_upload_large_document(self, api_base_url):
        """Test: Upload 10MB document should complete within reasonable time"""
        import time
        
        large_content = b"X" * (10 * 1024 * 1024)  # 10 MB
        payload = {
            "name": "large_file.bin",
            "mimeType": "application/octet-stream",
            "content": base64.b64encode(large_content).decode(),
            "createdBy": "PerfTest"
        }

        start_time = time.time()
        response = requests.post(f"{api_base_url}/api/documentstorage/upload", json=payload)
        duration = time.time() - start_time

        assert response.status_code == 200
        assert duration < 5.0  # Should complete in less than 5 seconds

    def test_retrieve_multiple_versions_performance(self, api_base_url):
        """Test: Retrieving version history with many versions should be fast"""
        import time

        # Create document with 20 versions
        upload_payload = {
            "name": "many_versions.txt",
            "mimeType": "text/plain",
            "content": base64.b64encode(b"V1").decode(),
            "createdBy": "PerfUser"
        }
        upload_response = requests.post(f"{api_base_url}/api/documentstorage/upload", json=upload_payload)
        master_id = upload_response.json()["masterId"]

        for i in range(2, 21):
            requests.post(f"{api_base_url}/api/documentstorage/{master_id}/save", json={
                "content": base64.b64encode(f"V{i}".encode()).decode(),
                "createdBy": "PerfUser"
            })

        # Measure retrieval time
        start_time = time.time()
        response = requests.get(f"{api_base_url}/api/documentstorage/{master_id}/versions")
        duration = time.time() - start_time

        assert response.status_code == 200
        assert len(response.json()) == 20
        assert duration < 1.0  # Should retrieve versions in less than 1 second
