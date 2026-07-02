import urllib.request
import urllib.parse
import json
import sys

BASE_URL = "http://127.0.0.1:5000"

# Use cookie processor to maintain session
opener = urllib.request.build_opener(urllib.request.HTTPCookieProcessor())

def request(path, method="GET", data=None):
    url = f"{BASE_URL}{path}"
    headers = {"Content-Type": "application/json"}
    
    req_data = None
    if data is not None:
        req_data = json.dumps(data).encode("utf-8")
        
    req = urllib.request.Request(url, data=req_data, headers=headers, method=method)
    try:
        with opener.open(req) as resp:
            body = resp.read().decode("utf-8")
            return resp.status, json.loads(body)
    except urllib.error.HTTPError as e:
        body = e.read().decode("utf-8")
        try:
            return e.code, json.loads(body)
        except:
            return e.code, body

def run_tests():
    print("Starting API validation tests...")
    
    # 1. Register Intern
    print("\n1. Testing /api/add_intern...")
    status, res = request("/api/add_intern", "POST", {
        "intern_id": "INT001",
        "intern_name": "John Doe",
        "department": "Engineering",
        "email": "john.doe@company.com"
    })
    print(f"Status: {status}, Response: {res}")
    assert status == 200 and res["success"] is True, "Add intern failed"

    # Register duplicate intern
    status, res = request("/api/add_intern", "POST", {
        "intern_id": "INT001",
        "intern_name": "Duplicate John",
        "department": "Engineering",
        "email": "john.doe@company.com"
    })
    print(f"Status: {status}, Response: {res}")
    assert status == 200 and res["success"] is False, "Add duplicate intern should fail"

    # 2. Get Interns
    print("\n2. Testing /api/interns...")
    status, res = request("/api/interns")
    print(f"Status: {status}, Response: {res}")
    assert status == 200 and len(res["interns"]) >= 1, "Get interns failed"
    assert res["interns"][0]["intern_id"] == "INT001", "Intern details mismatch"

    # 3. Intern Login
    print("\n3. Testing /api/login...")
    status, res = request("/api/login", "POST", {
        "intern_id": "INT001"
    })
    print(f"Status: {status}, Response: {res}")
    assert status == 200 and res["success"] is True, "Login failed"

    # Login again (should fail)
    status, res = request("/api/login", "POST", {
        "intern_id": "INT001"
    })
    print(f"Status: {status}, Response: {res}")
    assert status == 200 and res["success"] is False, "Login again should fail"

    # 4. Intern Logout
    print("\n4. Testing /api/logout...")
    status, res = request("/api/logout", "POST", {
        "intern_id": "INT001"
    })
    print(f"Status: {status}, Response: {res}")
    assert status == 200 and res["success"] is True, "Logout failed"

    # Logout again (should fail)
    status, res = request("/api/logout", "POST", {
        "intern_id": "INT001"
    })
    print(f"Status: {status}, Response: {res}")
    assert status == 200 and res["success"] is False, "Logout again should fail"

    # 5. Admin Auth
    print("\n5. Testing /api/admin/login...")
    status, res = request("/api/admin/me")
    print(f"Me before login: {res}")
    assert status == 200 and res["is_admin"] is False

    status, res = request("/api/admin/login", "POST", {
        "username": "admin",
        "password": "wrong_password"
    })
    print(f"Login wrong password status: {status}, Response: {res}")
    assert status == 200 and res["success"] is False

    status, res = request("/api/admin/login", "POST", {
        "username": "admin",
        "password": "admin123"
    })
    print(f"Login correct password status: {status}, Response: {res}")
    assert status == 200 and res["success"] is True

    status, res = request("/api/admin/me")
    print(f"Me after login: {res}")
    assert status == 200 and res["is_admin"] is True

    # 6. Apply Leave
    print("\n6. Testing /api/leave/apply...")
    status, res = request("/api/leave/apply", "POST", {
        "intern_id": "INT001",
        "leave_type": "Medical Leave",
        "from_date": "2026-06-08",
        "to_date": "2026-06-10",
        "reason": "Recovering from surgery"
    })
    print(f"Status: {status}, Response: {res}")
    assert status == 200 and res["success"] is True
    leave_id = res["leave_id"]

    # Get my leaves
    status, res = request(f"/api/leave/my_leaves?intern_id=INT001")
    print(f"Status: {status}, Response: {res}")
    assert status == 200 and len(res["leaves"]) >= 1

    # 7. Admin Review Leave
    print("\n7. Testing /api/leave/all_leaves and /api/leave/review...")
    status, res = request("/api/leave/all_leaves")
    print(f"All leaves status: {status}, Stats: {res.get('stats')}")
    assert status == 200 and len(res["leaves"]) >= 1

    status, res = request("/api/leave/review", "POST", {
        "leave_id": leave_id,
        "action": "Approved",
        "remarks": "Get well soon!"
    })
    print(f"Review status: {status}, Response: {res}")
    assert status == 200 and res["success"] is True

    # 8. Mark Absent (with Leave Awareness)
    print("\n8. Testing /api/mark_absent...")
    # Register a second intern who will be absent
    status, res = request("/api/add_intern", "POST", {
        "intern_id": "INT002",
        "intern_name": "Jane Doe",
        "department": "HR",
        "email": "jane.doe@company.com"
    })
    assert status == 200 and res["success"] is True

    # Run mark_absent for today
    status, res = request("/api/mark_absent", "POST", {})
    print(f"Status: {status}, Response: {res}")
    assert status == 200 and res["success"] is True
    # Jane Doe (INT002) should be marked absent
    # INT001 shouldn't be auto-marked absent because they logged in & out today

    # Test mark absent for the leave date (2026-06-09)
    status, res = request("/api/mark_absent", "POST", {"date": "2026-06-09"})
    print(f"Status for leave date: {status}, Response: {res}")
    assert status == 200 and res["success"] is True
    # INT001 should be marked "On Leave" because they have approved leave covering 2026-06-09
    assert "John Doe" in res["leave_marked"]

    # 9. Dashboard check
    print("\n9. Testing /api/dashboard...")
    status, res = request("/api/dashboard")
    print(f"Status: {status}, Today's Stats: {res['today']}")
    assert status == 200
    assert res["today"]["total_registered"] >= 2

    # 10. Admin Logout
    print("\n10. Testing /api/admin/logout...")
    status, res = request("/api/admin/logout", "POST", {})
    print(f"Status: {status}, Response: {res}")
    assert status == 200 and res["success"] is True

    status, res = request("/api/admin/me")
    assert status == 200 and res["is_admin"] is False

    print("\nALL API TESTS PASSED SUCCESSFULLY!")

if __name__ == "__main__":
    try:
        run_tests()
    except AssertionError as e:
        print(f"\nTEST FAILED: {e}", file=sys.stderr)
        sys.exit(1)
    except Exception as e:
        print(f"\nUNEXPECTED ERROR: {e}", file=sys.stderr)
        sys.exit(1)
