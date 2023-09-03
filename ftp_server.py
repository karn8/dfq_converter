from pyftpdlib.authorizers import DummyAuthorizer
from pyftpdlib.handlers import FTPHandler
from pyftpdlib.servers import FTPServer

# Define the server settings
server_ip = "127.0.0.1"
server_port = 2121
username = "user"
password = "password"

def setup_ftp_server():
    authorizer = DummyAuthorizer()
    authorizer.add_user(username, password, ".", perm="elradfmw")  # Modify permissions as needed

    handler = FTPHandler
    handler.authorizer = authorizer

    server = FTPServer((server_ip, server_port), handler)
    print(f"FTP server started on {server_ip}:{server_port}")
    server.serve_forever()

if __name__ == "__main__":
    setup_ftp_server()
