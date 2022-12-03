import http.server
import socketserver
import sys

PORT = 8000


def main():
    global PORT
    _port = PORT
    if len(sys.argv) == 2:
        try:
            _port = int(sys.argv[1])
        except Exception:
            _port = PORT

    Handler = http.server.SimpleHTTPRequestHandler
    with socketserver.TCPServer(("", _port), Handler) as httpd:
        print("serving at port", _port)
        httpd.serve_forever()


if __name__ == "__main__":
    main()
