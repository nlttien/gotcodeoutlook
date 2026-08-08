"""
Script kiểm thử API Server (/api/get-code)
"""

import sys
from fastapi.testclient import TestClient

try:
    from api_server import app
except ModuleNotFoundError as err:
    print(f"[!] Lỗi import: {err}")
    sys.exit(1)

client = TestClient(app)


def test_health_endpoint():
    print("=" * 60)
    print(" 1. KIỂM TRA ENDPOINT GET /health")
    print("=" * 60)
    res = client.get("/health")
    print(f"Status Code: {res.status_code}")
    print(f"Response   : {res.json()}")
    assert res.status_code == 200


def test_get_code_mock():
    print("\n" + "=" * 60)
    print(" 2. KIỂM TRA ENDPOINT POST /api/get-code (MOCK MODE)")
    print("=" * 60)
    
    account_str = (
        "christinakazunas1125@outlook.com|"
        "ChristinaKazunas694|"
        "M.C542_BAY...|"
        "9e5f94bc-e8a4-4e73-b8be-63364c29d753"
    )
    
    payload = {
        "account_str": account_str,
        "keyword": "mã xác nhận",
        "use_mock": True
    }
    
    res = client.post("/api/get-code", json=payload)
    print(f"Status Code: {res.status_code}")
    data = res.json()
    print("Response JSON:")
    for k, v in data.items():
        print(f"  {k:20s}: {v}")
        
    assert res.status_code == 200
    assert data["status"] == "success"
    assert data["otp_code"] == "849201"
    print("\n -> KẾT QUẢ: TRÍCH XUẤT MÃ OTP '849201' THÀNH CÔNG!")


def test_get_code_real():
    print("\n" + "=" * 60)
    print(" 3. KIỂM TRA ENDPOINT POST /api/get-code (REAL ACCOUNT CONNECTION)")
    print("=" * 60)
    
    account_str = (
        "devinecrystal9187@outlook.com|"
        "DevineCrystal37720|"
        "M.C517_BAY.0.U.MsaArtifacts.-ClVMh8nbP*ZTBY4BsXv7iRI3gcrb8d1Hai1cp6HANOlcYGrqz6YcgsdygC*6kpqs0pxlHWU5DNmq*pNFhpiX72DgxyNz*m3Eox2547bwmXbYEZF9ZuExZLvn1NXeEPEQsfyGnBXZC*VY2Q4!GExqrZ87udvjAzMfr1Af3Epc6DgROw8CgLcNVJbuokrD!0zthfeNZY2mNbKfU0tLt5x!FkTR2h0jnjpeEShcyKXESGCMLIIMOEYpIpKm4*c0Nn3wbX2*jg!sddPx8gjWeY!Sb77Qo4XtvS6VhXAHSAEmJhHQFE6cI9gic7TUZSUl660F3jNdC8PmTMxLcpwuYr6crsFVLFxBAegUJW3SgW*mii8kJYPbk!g013xiKGzw5oC!1eJtGEPgdy*9BVdyZqlI6r4$|"
        "9e5f94bc-e8a4-4e73-b8be-63364c29d753"
    )
    
    payload = {
        "account_str": account_str,
        "use_mock": False
    }
    
    res = client.post("/api/get-code", json=payload)
    print(f"Status Code: {res.status_code}")
    data = res.json()
    print("Response JSON:")
    for k, v in data.items():
        print(f"  {k:20s}: {v}")
        
    assert res.status_code == 200
    print("\n -> KẾT QUẢ: KẾT NỐI VÀ ĐỌC HÒM THƯ THẬT THÀNH CÔNG!")


if __name__ == '__main__':
    test_health_endpoint()
    test_get_code_mock()
    test_get_code_real()

