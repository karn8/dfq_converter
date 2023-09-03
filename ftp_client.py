from ftplib import FTP

def upload_file(filename, server_ip, server_port, username, password):
    with FTP() as ftp:
        ftp.connect(server_ip, server_port)
        ftp.login(username, password)
        with open(filename, "rb") as file:
            ftp.storbinary(f"STOR {filename}", file)
        print(f"File '{filename}' uploaded successfully.")

if __name__ == "__main__":
    server_ip = "127.0.0.1"
    server_port = 2121
    username = "user"
    password = "password"
    file_to_upload = "file_to_upload.txt"

    upload_file(file_to_upload, server_ip, server_port, username, password)
