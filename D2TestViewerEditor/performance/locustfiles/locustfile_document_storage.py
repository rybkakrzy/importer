"""
Locust performance test for Document Storage with Versioning
Run: locust -f locustfile_document_storage.py --host=http://localhost:5190
"""
from locust import HttpUser, task, between, constant_pacing
import base64
import random
import string


class DocumentStorageUser(HttpUser):
    """
    Simulates user uploading, editing, and restoring document versions
    """
    wait_time = between(1, 3)  # Wait 1-3 seconds between tasks
    
    def on_start(self):
        """Initialize user session"""
        self.master_ids = []
        self.version_history = {}  # {master_id: [version_ids]}
        
    def generate_random_content(self, size_kb=10):
        """Generate random binary content"""
        size_bytes = size_kb * 1024
        content = ''.join(random.choices(string.ascii_letters + string.digits, k=size_bytes))
        return base64.b64encode(content.encode()).decode()

    @task(5)
    def upload_new_document(self):
        """Upload new document (weighted 5x - most common operation)"""
        payload = {
            "name": f"document_{random.randint(1000, 9999)}.txt",
            "mimeType": "text/plain",
            "content": self.generate_random_content(size_kb=5),
            "createdBy": f"User{self.client.user_id if hasattr(self.client, 'user_id') else 'Test'}"
        }
        
        with self.client.post("/api/documentstorage/upload", json=payload, catch_response=True) as response:
            if response.status_code == 200:
                result = response.json()
                master_id = result["masterId"]
                version_id = result["versionId"]
                
                self.master_ids.append(master_id)
                self.version_history[master_id] = [version_id]
                
                response.success()
            else:
                response.failure(f"Upload failed: {response.status_code}")

    @task(8)
    def save_document_version(self):
        """Save new version of existing document (weighted 8x - most frequent)"""
        if not self.master_ids:
            return  # Skip if no documents uploaded yet
        
        master_id = random.choice(self.master_ids)
        payload = {
            "content": self.generate_random_content(size_kb=random.randint(5, 15)),
            "createdBy": f"Editor{random.randint(1, 5)}"
        }
        
        with self.client.post(f"/api/documentstorage/{master_id}/save", json=payload, catch_response=True) as response:
            if response.status_code == 200:
                result = response.json()
                version_id = result["versionId"]
                
                if master_id in self.version_history:
                    self.version_history[master_id].append(version_id)
                
                response.success()
            else:
                response.failure(f"Save version failed: {response.status_code}")

    @task(3)
    def get_active_document(self):
        """Retrieve active version of document (weighted 3x)"""
        if not self.master_ids:
            return
        
        master_id = random.choice(self.master_ids)
        
        with self.client.get(f"/api/documentstorage/{master_id}", catch_response=True) as response:
            if response.status_code == 200:
                document = response.json()
                assert "content" in document
                assert "versionNumber" in document
                response.success()
            else:
                response.failure(f"Get document failed: {response.status_code}")

    @task(2)
    def get_version_history(self):
        """Get all versions (version history UI) (weighted 2x)"""
        if not self.master_ids:
            return
        
        master_id = random.choice(self.master_ids)
        
        with self.client.get(f"/api/documentstorage/{master_id}/versions", catch_response=True) as response:
            if response.status_code == 200:
                versions = response.json()
                assert isinstance(versions, list)
                
                # Verify sorting
                if len(versions) > 1:
                    for i in range(len(versions) - 1):
                        assert versions[i]["versionNumber"] >= versions[i+1]["versionNumber"]
                
                response.success()
            else:
                response.failure(f"Get versions failed: {response.status_code}")

    @task(1)
    def restore_previous_version(self):
        """Restore old version (weighted 1x - least common)"""
        if not self.master_ids:
            return
        
        master_id = random.choice(self.master_ids)
        
        # Get versions first
        if master_id not in self.version_history or len(self.version_history[master_id]) < 2:
            return  # Need at least 2 versions to restore
        
        # Pick random old version (not the latest)
        old_version_id = random.choice(self.version_history[master_id][:-1])
        
        with self.client.post(f"/api/documentstorage/{master_id}/restore/{old_version_id}", json={}, catch_response=True) as response:
            if response.status_code == 200:
                response.success()
            else:
                response.failure(f"Restore version failed: {response.status_code}")


class DocumentStorageHeavyUser(HttpUser):
    """
    Heavy user scenario - uploads large documents
    """
    wait_time = constant_pacing(5)  # One request every 5 seconds
    
    def generate_large_content(self, size_mb=1):
        """Generate large binary content"""
        size_bytes = size_mb * 1024 * 1024
        content = bytes(random.getrandbits(8) for _ in range(size_bytes))
        return base64.b64encode(content).decode()

    @task
    def upload_large_document(self):
        """Upload 1-5 MB document"""
        size_mb = random.randint(1, 5)
        payload = {
            "name": f"large_file_{size_mb}MB.bin",
            "mimeType": "application/octet-stream",
            "content": self.generate_large_content(size_mb=size_mb),
            "createdBy": "HeavyUser"
        }
        
        with self.client.post("/api/documentstorage/upload", json=payload, catch_response=True, name="/api/documentstorage/upload [LARGE]") as response:
            if response.status_code == 200:
                response.success()
            else:
                response.failure(f"Large upload failed: {response.status_code}")


class DocumentStorageReadHeavyUser(HttpUser):
    """
    Read-heavy user - mostly retrieves documents
    """
    wait_time = between(0.5, 2)
    
    def on_start(self):
        """Setup: upload some documents for reading"""
        self.master_ids = []
        
        for i in range(5):
            payload = {
                "name": f"readonly_doc_{i}.txt",
                "mimeType": "text/plain",
                "content": base64.b64encode(f"Content {i}".encode()).decode(),
                "createdBy": "SetupUser"
            }
            response = self.client.post("/api/documentstorage/upload", json=payload)
            if response.status_code == 200:
                self.master_ids.append(response.json()["masterId"])

    @task(10)
    def read_document(self):
        """Read documents frequently"""
        if not self.master_ids:
            return
        
        master_id = random.choice(self.master_ids)
        self.client.get(f"/api/documentstorage/{master_id}")

    @task(5)
    def read_versions(self):
        """Read version history"""
        if not self.master_ids:
            return
        
        master_id = random.choice(self.master_ids)
        self.client.get(f"/api/documentstorage/{master_id}/versions")


class DocumentStorageVersioningStressTest(HttpUser):
    """
    Stress test - creates many versions rapidly
    """
    wait_time = constant_pacing(0.5)  # 2 requests per second per user
    
    def on_start(self):
        """Upload initial document"""
        payload = {
            "name": "stress_test.txt",
            "mimeType": "text/plain",
            "content": base64.b64encode(b"Initial").decode(),
            "createdBy": "StressUser"
        }
        response = self.client.post("/api/documentstorage/upload", json=payload)
        if response.status_code == 200:
            self.master_id = response.json()["masterId"]
        else:
            self.master_id = None

    @task
    def rapid_version_creation(self):
        """Create versions rapidly to stress test database"""
        if not hasattr(self, 'master_id') or not self.master_id:
            return
        
        payload = {
            "content": base64.b64encode(f"Version {random.randint(1, 10000)}".encode()).decode(),
            "createdBy": f"StressUser{random.randint(1, 100)}"
        }
        
        with self.client.post(f"/api/documentstorage/{self.master_id}/save", json=payload, name="/api/documentstorage/[master]/save [STRESS]") as response:
            if response.status_code != 200:
                print(f"Stress test failure: {response.status_code} - {response.text}")
