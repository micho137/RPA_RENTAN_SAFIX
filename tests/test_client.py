from src.outlook.client import OutlookClient

def test_client_has_methods():
    print("\n[TEST] Verificando que OutlookClient tenga los métodos esperados...")
    c = OutlookClient  # No lo instanciamos para no tocar COM real
    expected = ["find_store_by_display", "get_folder", "iter_items", "mark_as_read", "move_to", "attachments"]

    for method in expected:
        print(f"  - Verificando método: {method}")
        assert hasattr(c, method), f"Falta el método: {method}"

    print("[OK] Todos los métodos existen correctamente en OutlookClient ✅")
