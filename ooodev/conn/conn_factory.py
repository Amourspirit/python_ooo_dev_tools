from __future__ import annotations
import json
from ooodev.conn import connectors


def get_from_json(s: str = "") -> connectors.ConnectorBridgeBase | None:
    """
    Gets a connector from a json string.

    Args:
        s (str, optional): Json str. Defaults to "".

    Raises:
        ValueError: If no connector key found in json data
        ValueError: If unknown connector key found in json data

    Returns:
        connectors.ConnectorBridgeBase | None: Connector if found, else ``None``

    Note:
        If ``None`` is returned this usually means that the connection is a direct local connection as in a macro.

        The serialize string is usually gotten from the ``connector.serialize()`` method.

    .. versionadded:: 0.44.0
    """
    if not s:
        return None
    data = json.loads(s)
    if "connector" not in data:
        raise ValueError("No connector key found in json data")
    key = data["connector"]
    if key == "pipe":
        return connectors.ConnectPipe.deserialize(s)
    if key == "socket":
        return connectors.ConnectSocket.deserialize(s)
    raise ValueError(f"Unknown connector key: {key}")
