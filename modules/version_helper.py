import AutoUpdate

class VersionHelper():
    def __init__(self) -> None:
        AutoUpdate.set_url("https://raw.githubusercontent.com/FlorianZasa/capza/main/remote_version.txt?token=GHSAT0AAAAAABZNEWIL4SRXREBBLVY63JF6Y3EG5SA")
        AutoUpdate.set_current_version("0.1.2")

    def run():
        if not AutoUpdate.is_up_to_date():
            AutoUpdate.download("")

        
                


if __name__ == "__main__":
    dh = VersionHelper()
