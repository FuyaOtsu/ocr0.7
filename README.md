# ocr0.7

# ゲーム攻略補助デスクトップアプリ「ocr」

ゲーム画面を画像認識し、任務の攻略情報を提示するデスクトップアプリケーションです。
本バージョンは試作段階であり、近日製品版にアップデートさせていただきます。
「ocr0.7.zip」をダウンロードしていただき、解凍後するだけで使用可能です。

# デモ

![アプリ動作デモ](https://github.com/FuyaOtsu/ocr0.7/blob/main/ocr_gif.gif)
Gifでは一つの任務ですが、複数の任務を選択可能です。

# 製作理由

艦これの任務名はなかなか一発で変換しづらいものが多く、検索するときに手間がかかるため。

# 主な機能
・艦これの画面をキャプチャすることで、画像認識を用いて任務名を認識します。

・認識した任務名をgoogle検索にかけ、攻略に必要な情報を抽出し、自動で表示させます。

・検索し、情報を抽出させて頂いたサイト様の規約に違反しないよう、最新の注意を払っています。（万が一何かあれば、「はじめに」内の連絡先へ）

・艦これ側のデータの送受信に、一切影響しません。（垢BANのリスクが極めて低い）※使用はあくまでも自己責任でお願いします。

# 力を入れた点や今後の課題
・現状の艦これ補助ソフトは、艦これがサーバーから受信したデータを使用するものが多く、それに対し忌避感や、うしろめたさを感じる人が多い現状です。それに対し私は、画像認識の専門性を活かし、より多くの人が安心して使っていただけるソフトの制作を試みました。

・攻略情報をただ出力するだけでなく、出撃先によってタブを分けるなど、視覚的な使用感にもこだわりました。

・今後は、より多くのプレイヤーに使用していただけるよう、画像認識の精度向上に努めます。

# その他
・使用時には、付属の「はじめに」をお読みください。

・本ソースコードは、pythonを使用しています。依存ライブラリ等は、「rib」に記載しています。


