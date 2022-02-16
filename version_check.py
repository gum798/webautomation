from win32com.client import Dispatch
import urllib.request

def get_version_via_com(filename):
    parser = Dispatch("Scripting.FileSystemObject")
    try:
        version = parser.GetFileVersion(filename)
    except Exception:
        return None
    return version

def get_chrome_version():
    paths = [r"C:\Program Files\Google\Chrome\Application\chrome.exe",
             r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"]
    version = list(filter(None, [get_version_via_com(p) for p in paths]))[0]
    return version


def get_last_version():
    url = "https://chromedriver.storage.googleapis.com/LATEST_RELEASE"
    import requests
    response = requests.get(url)
    return response.text

def unzip():
    import zipfile
    zip_ref = zipfile.ZipFile("chromedriver_win32.zip", 'r')
    zip_ref.extractall(".")
    zip_ref.close()
    print("Unzipped the latest version of chromedriver")
        
def download_file(version):
    urllib.request.urlretrieve("https://chromedriver.storage.googleapis.com/index.html?path="+version, "chromedriver_win32.zip")
    print("Downloaded the latest version of chromedriver")
    unzip()



if __name__ == "__main__":
    version = get_chrome_version()
    #설치된 크롬 버전 확인
    print(version)
    print(version.split(".")[0])
    #크롬드라이버 버전 확인
    
    #최신 크롬드라이버 버전 확인
    version = get_last_version()
    print(version)
    print(version.split(".")[0])
