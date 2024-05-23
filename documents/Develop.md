# Develop

## 環境
- Windows 11
- Python 3.11.3

## 環境構築
```powershell
# Clone
$ git clone https://github.com/TaichiOikawa/Mosaic-Art/

$ cd Mosaic-Art

# venv環境の構築
$ python3 -m venv venv

# 仮想環境を起動 (Powershell)
$ venv\Scripts\Activate.ps1
# 仮想環境を起動 (Mac)
$ source venv/bin/activate

# ライブラリをインストール
$ pip install -r requirements.txt
```

## 実行
```powershell
$ python3 main.py
```

## EXE化
```powershell
# どちらも仮想環境下で行う
# pyinstallerをインストール
$ pip install pyinstaller

# exe化
$ pyinstaller main.py --onefile --clean
```


## メモ
### openpyxlの仕様
- 行幅、列幅、余白等、だいたいずれます。根気強く調整すること。
- なぜかページサイズにB4はありません。あと、Excel上で指定する場合は「JIS B4 (25.71*36.4)」を使用すること。
