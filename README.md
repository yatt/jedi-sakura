# jedi-sakura

## これは何？
jedi-sakura は、Pythonの静的なコード自動補完ライブラリ jedi ( http://jedi.jedidjah.ch/en/latest/ )
の、サクラエディタ( http://sakura-editor.sourceforge.net/ )用マクロです。

jedi-sakuraを使う事で、サクラエディタのPython用エディタとしての能力を強化できます。

![screenshot](https://github.com/yatt/jedi-sakura/raw/master/docs/screenshots/sample-animation.gif)


## できること
* サクラエディタで、jedi-vimライクな補完ができる。

## できないこと
* サクラエディタの現在の入力補完の仕様上、補完は先頭一致のみサポート
* サクラエディタの現在の入力補完の仕様上、スペース直後では補完不能（「import 」からモジュール名を補完するようなのは無理）

## インストール要件
* python
* jedi ( https://pypi.python.org/pypi/jedi/ )
* サクラエディタ 2.0.6 or higher

## インストール手順
通常のサクラエディタプラグインと同様。

以上


