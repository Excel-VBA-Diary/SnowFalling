# SnowFalling<br>

機能<br>
エクセルのワークシートに雪を降らせる（Snow falling on Excel worksheet)<br>
基本構造は雪のクラスとコレクション、ラウンドロビン方式で一つひとつの雪を動かし最後に消している。<br>

1.準備<br>
(1) 標準モジュール Module1.bas をインポートする<br>
(2) クラスモジュール Snow.cls をインポートする<br>
(3) 任意の名前でマクロブックを保存する<br>
(4) マクロブックを保存したフォルダーにsnowサブフォルダーを作り、Snow.zip を解凍して画像を格納する<br>

2.起動方法<br>
Mainプロシージャを起動する。アクティブシートにボタンを配置すると便利。<br>
