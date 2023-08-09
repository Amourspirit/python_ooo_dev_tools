# TEST Notes

Test are all written for pytest.

## Environment Variables

A `.test.env` file can be create in the project root to set environment variables for testing. This file is ignored by git.

Default values

```ini
ODEV_TEST_HEADLESS=1
ODEV_TEST_OPT_DYNAMIC=0
ODEV_TEST_OPT_VISIBLE=0
ODEV_TEST_CONN_SOCKET=0
ODEV_TEST_CONN_SOCKET_PORT=2002
ODEV_TEST_CONN_SOCKET_HOST=localhost
ODEV_TEST_CONN_SOCKET_KIND=default
ODEV_TEST_OPT_VERBOSE=1
```

- `ODEV_TEST_OPT_DYNAMIC` - If set to `1` then the tests will run in dynamic mode. This is a property of the `Options` class. In short, dynamic mode will get dynamic Values for `Lo.this_component` and `Lo.xcomponent_context`.
- `ODEV_TEST_OPT_VISIBLE` - If set to `1` then the tests will start LibreOffice with the `--invisible` flag set to `false`.
- `ODEV_TEST_CONN_SOCKET ` - If set to `1` then the tests will use a socket connection to LibreOffice; Otherwise, the connection is made via pipe.
- `ODEV_TEST_CONN_SOCKET_KIND` = `default` or `no_start`, if `no_start` then the connection will not start LibreOffice.
