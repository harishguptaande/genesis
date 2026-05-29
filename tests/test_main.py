def test_read_root(client):
    response = client.get("/")
    assert response.status_code == 200
    assert "message" in response.json()


def test_health_check(client):
    response = client.get("/health")
    assert response.status_code == 200
    assert response.json()["status"] == "healthy"


def test_create_item(client):
    response = client.post(
        "/items/",
        json={"name": "Test Item", "description": "A test item"}
    )
    assert response.status_code == 200
    assert response.json()["name"] == "Test Item"
    assert "id" in response.json()


def test_list_items(client):
    client.post("/items/", json={"name": "Item 1"})
    client.post("/items/", json={"name": "Item 2"})
    
    response = client.get("/items/")
    assert response.status_code == 200
    assert len(response.json()) == 2


def test_get_item(client):
    create_response = client.post("/items/", json={"name": "Test Item"})
    item_id = create_response.json()["id"]
    
    response = client.get(f"/items/{item_id}")
    assert response.status_code == 200
    assert response.json()["name"] == "Test Item"


def test_update_item(client):
    create_response = client.post("/items/", json={"name": "Original"})
    item_id = create_response.json()["id"]
    
    response = client.put(
        f"/items/{item_id}",
        json={"name": "Updated", "completed": True}
    )
    assert response.status_code == 200
    assert response.json()["name"] == "Updated"
    assert response.json()["completed"] is True


def test_delete_item(client):
    create_response = client.post("/items/", json={"name": "To Delete"})
    item_id = create_response.json()["id"]
    
    response = client.delete(f"/items/{item_id}")
    assert response.status_code == 200
    
    get_response = client.get(f"/items/{item_id}")
    assert get_response.status_code == 404
