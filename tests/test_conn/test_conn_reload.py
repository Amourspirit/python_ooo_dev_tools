import pytest

if __name__ == "__main__":
    pytest.main([__file__])


def test_conn_reload(loader) -> None:
    """
    Version 0.53.0 introduced ConnectCtx and a new option to force reload of LibreOffice connection.
    This test ensures that the option works.
    ConnectCtx is pass the ctx to create the connection.
    In this test we must use the current context to create the connection.
    In An extension you might use the ctx of the extension method.
    """
    from ooodev.conn.connect_ctx import ConnectCtx
    from ooodev.loader.lo import Lo, _on_global_document_event

    opt = Lo.Options(force_reload=True)
    inst = Lo.current_lo
    Lo.load_office(connector=ConnectCtx(inst.get_context()), opt=opt)
    assert Lo._lo_inst is not None
    assert Lo._lo_inst is not inst
    inst.global_event_broadcaster.add_event_document_event_occurred(_on_global_document_event)
    Lo._lo_inst = inst
