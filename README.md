# CSVインポート/エクスポート
python: 3.7.1 以上  
poetry: 1.1.4 以上

## パッケージ管理
`poetry`を使用  
実行環境だけなら`poetry`なしでも`pip install`できるらしい。http://orolog.hatenablog.jp/entry/2019/03/24/223531

### Poetryインストール
基本的に左記URL参照: https://cocoatomo.github.io/poetry-ja/

※`bash`や`zsh`の補完の有効化は追加手順があることに注意

#### Poetry設定変更
`.venv`をプロジェクトのディレクトリ内に作成する設定

```sh
poetry config virtualenvs.in-project true
```

参考: https://qiita.com/ragnar1904/items/0e5b8382757ccad9a56c

##### VSCodeからの参照
ワークスペースのディレクトリ配下はデフォルトで読んでくれる(後述のコマンドで実行できるので不要かも)  
参考: https://tekunabe.hatenablog.jp/entry/2018/12/28/vscode_venv_default_rolad

### Poetryざっくり使い方
`pyproject.toml`のファイルで依存関係が管理されていて、`poetry.lock`でバージョンが固定される。

参考: https://cocoatomo.github.io/poetry-ja/basic-usage/  
コマンド: https://cocoatomo.github.io/poetry-ja/cli/

#### パッケージインストール
下記コマンドで`.venv`環境にパッケージをインストール

```sh
poetry install
```

##### 最新パッケージの取得
`poetry.lock`を更新して最新バージョンのパッケージを取得

```sh
poetry update
```

#### パッケージ追加
下記コマンドで`pyproject.toml`にパッケージを追加してインストール

```sh
poetry add <packages> ..
```

##### パッケージ削除

```sh
poetry remove <packages> ..
```

#### 実行
`.venv`の仮想環境の中でコマンド実行

```sh
poetry run python sample.py
```

`.venv`の仮想環境の中でシェル起動(終了は`exit`)

```sh
poetry shell
```

## CSV, Excelファイル操作
`pandas`と`openpyxl`というパッケージを使用する予定。

TODO: 必要になったら追記する  
`pandas`: データの整形がやりやすくなる予定  
`openpyxl`: Excelファイルの操作が可能

## DBアクセス
`psycopg2`: PostgreSQLドライバ


## ローカル開発環境(WSL2)
ストアから取れるUbuntu使う。

### インストール
公式ドキュメントそのまま。Ubuntu入れた。Windowsターミナルも便利そうだから入れた
https://docs.microsoft.com/ja-jp/windows/wsl/install-win10

VSCodeの拡張機能もつよい(仮想マシンからVSCodeひらける)
https://qiita.com/EBIHARA_kenji/items/12c7a452429d79006450

#### 機能の有効化関連
手順の中で↓2つを有効化してる
* Linux用Windowsサブシステム
* 仮想マシンプラットフォーム
これはコントロールパネル->プログラムと機能->Windowsの機能の有効化または無効化でも設定可能
(他の仮想化プラットフォームと競合したりしたら)

### 環境コピー
インストール直後の環境をコピーする
https://qiita.com/souyakuchan/items/9f95043cf9c4eda2e1cc

```
> wsl --export Ubuntu Ubuntu-20.04_copy.tar
> wsl --import Ubuntu_app .\wsl_manual_install\ubuntu2004\app\ Ubuntu-20.04_copy.tar
> wsl --import Ubuntu_postgres .\wsl_manual_install\ubuntu2004\postgres\ Ubuntu-20.04_copy.tar
> wsl --import Ubuntu_sqlserver .\wsl_manual_install\ubuntu2004\sqlserver\ Ubuntu-20.04_copy.tar
```

`.\wsl_manual_install\ubuntu2004\..` はインストール先のパス(環境ごとにフォルダ分けないといけないっぽい)

### 接続
Windows Terminalのプロファイルに自動で追加されるので接続は簡単(デフォルトの設定でそうなってるはず)
https://docs.microsoft.com/ja-jp/windows/wsl/install-win10#install-windows-terminal-optional

Windows Terminal使わないなら `wsl -d <環境名>` で接続
カレントディレクトリはWindowsそのまま(`/mnt/c/..`以下にCドライブみたいな感じでマウントされる)
ユーザは`root`で入るので`su - <ユーザ名>`等で適宜ユーザ切り替え

### 環境削除

```
> wsl --unregister <環境名>
```