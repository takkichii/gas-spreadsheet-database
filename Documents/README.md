# 概要
こちらの記事では、GAS（Google Apps Script）で Google スプレッドシートを Database のように直感的に操作する方法をご紹介します。

簡単に言うと、<strong>ヘッダー（1行目の項目名）を定義</strong>しさすえれば、そのヘッダーの項目名をキーとしてデータを検索したり、追加、更新、削除ができるようになります。

# メリット
スプレッドシートのヘッダー（1行目）を基点にデータ操作を行うので、行列の構成を変更しても影響を受けません。
例えば、列の順番を入れ替えてもデータ操作に影響はありませんし、列の前後に新しい項目名の列を追加しても影響はありません。

こちらのテンプレートを使用することで、```getRange()``` などの<strong><font color="#FF6666">セルレベルで指定していたデータ操作の煩わしさから解放されます。</font></strong>

# 活用例
こちらの機能を土台にアプリケーションを開発しました。
開発例としてご参考ください。

- [Qiita - タイムズカーシェアの予約をカレンダーに自動反映する](https://qiita.com/takkichii/items/1a262cad1fe35c72b6a2)
- [Qiita - GAS でアプリを作る【面接評価シート】](https://qiita.com/BruceWeyne/items/546491b280d2e42123ca)

# はじめに
GAS（Google Apps Script）を時折作成する中で、スプレドシートのデータを操作したい場合があります。
公式ドキュメントや記事を読みながら ```SpreadsheetApp``` クラスや ```getRange()```, ```getValues()``` メソッドなどの活用が必要になりますが、使い方を忘れてしまい調べ直すところから取り組むのが常でした。
それがとても煩わしく、データ操作の仕方に惑わされることなく気軽にアプリケーション開発がしたいな、と思い始め、思い切ってベース機能を開発することにしました。

# 目的
開発の目的は GAS からスプレッドシートのデータ操作を行う際に、毎度ドキュメントを読み返す必要なく直感的に操作できる仕組みを開発することです。
ここでの「データ操作」とは、データの「取得・登録・更新・削除」を意味します。

# こんな方々へ
GAS は法人アカウントでも無料アカウントでも使える Google の素晴らしいサービスです。
業務効率化や独自サービスのために既にアプリケーションを開発されていたり、これから勉強がてら触れてみようと思われている方もいると思います。
色んな目的があると思いますが、こちらの記事を見つけてくださった方々の手間を省く便利機能として役立ってもらえたら嬉しいです。

# ファイル構成
以下が開発したテンプレートファイルの一覧です。

```javascript
- Config.gs     // 操作するスプレッドシートの ID や他の初期設定の定数類を管理する
- Controller.gs // ここでアプリケーションのコードを自由に書きます（サンプル）
- Database.gs   // スプレドシート（Database）を実際に操作する
- Model.gs      // Database ファイルを読み込み（インスタンス化）します
```
:::note info
<strong>重要なのは Config, Database, Model ファイルの 3 つ</strong>で、Controller ファイルは飽くまでサンプルなので使用しなくても大丈夫です。
:::

## ソースコード
■ソースコードは <strong>[こちら](https://github.com/BruceWeyne/gas-spreadsheet-database) </strong>の GitHub に格納しています。

# 使い方
## 実装方法
### GASの初期設定
実装は簡単で、上記のソースコードをそのままご自身の GAS 環境にコピペしてください。
ファイル名は全く一緒にしなくても大丈夫ですが、Javascript のクラスを活用しているので、クラス名とファイル名を合わせていただくと理解しやすいかと思います。

![gas-implement-sample.png](https://qiita-image-store.s3.ap-northeast-1.amazonaws.com/0/364477/b1d4bae9-1810-8038-00fd-7379a510b8e8.png)

### スプレッドシートの用意
データ操作に使用したいスプレッドシートをご自身の Google ドライブに用意してください。
その URL に記載のスプレッドシート ID を Config ファイルの ```spreadsheetId``` 変数に設定します。

![spreadsheet-sample.png](https://qiita-image-store.s3.ap-northeast-1.amazonaws.com/0/364477/dd5e6a07-3505-5e6c-f840-46d72ea97d08.png)

![config-sample.png](https://qiita-image-store.s3.ap-northeast-1.amazonaws.com/0/364477/9f25caca-97ec-cf3a-605a-948071adeb4b.png)

:::note
スプレッドシート ID を設定しない場合は ```getActiveSpreadsheet()``` として機能します。
従って、スプレッドシート内のプロジェクトとして GAS ファイルを作成した場合は ```spreadsheetId``` に ID を設定しなくても問題ありません。
:::

### スプレッドシートにヘッダーを設定する
スプレッドシートを Database ライクに操作するので、ヘッダー（1行目の項目名）は必須です。
また、シート名もデータ操作の要になります。

関連イメージとしては、

| スプレッドシート | Database  | 
|:-----------|:------------|
| シート（名）    | Table（名）   | 
| シートの1行目 | Header（項目名） | 
| 2行目以降   | データ       |

という位置付けです。

![sheet-sample.png](https://qiita-image-store.s3.ap-northeast-1.amazonaws.com/0/364477/f246d16c-d268-60e8-8f18-ce6ace61fbde.png)

## データ操作の例
これで初期設定は完了しました。
あとはご自身の好きなアプリケーションを作るまでです。
ここからは、データ操作の一例をお見せします。

### Modelインスタンスの生成
まず、データ操作をするために Model クラスのインスタンスを生成します。

```javascript:インスタンスの生成
const mdl = new Model();
```

これで Config ファイルに設定したスプレッドシートがデフォルトとして読み込まれます。
もし他のスプレッドシートのデータ操作を行いたい場合は、

```javascript:他のスプレッドシートの読み込み
const spreadsheetId2 = "10LA-e_vJuFcN13HIwuRoJPZTehto17Z9lig-c424Ig0";
const mdl2 = new Model(spreadsheetId2);
```

と Model クラス引数にスプレッドシートの ID を設定することもできます。

:::note warn
スプレッドシートの ID は Config ファイルで管理していただくことを推奨します。
:::

### データの取得
それでは、実際にデータを読み込んでみます。
Controller ファイルにアプリケーションのコードを記述していきますが、飽くまでサンプルなので Controller ファイルではなく、ご自身のお好きなファイルに記述してください。

データの取得には ```getData()``` メソッドを使用します。

```javascript:getData
getData(sheetName, conditions, offsetRow, limitRow)

[arg1] sheetName  : シート名
[arg2] conditions : 絞り込み条件
[arg3] offsetRow  : データ取得の開始数
[arg4] limitRow   : データ取得最大数
```

```javascript:データの全取得
function myFunction() {
    // インスタンスの生成
    const mdl = new Model();
    
    // データの全取得
    const data = mdl.getData("応募者");
    
    // ログ出力
    Logger.log(data);
    Logger.log("===============");
    Logger.log(data[0]["氏名"]);
    Logger.log(data[0]["年齢"]);
    Logger.log(data[0]["性別"]);
}
```
上記を実行すると以下のようになります。
データをオブジェクトの連想配列のリストで取得していることが分かると思います。

![data-get-sample.png](https://qiita-image-store.s3.ap-northeast-1.amazonaws.com/0/364477/e3f22fc3-dfd2-b023-3d22-38c824a57b7f.png)

#### データのAND検索
さらに、データ取得時に絞り込みをすることも可能です。

```javascript:データの AND 検索取得
function myFunction() {
    // インスタンスの生成
    const mdl = new Model();
    
    // 検索条件を設定
    const conditions = [
      { key: "氏名", value: "天野 結女" },
      { key: "年齢", value: 21 }
    ];
    
    // データの AND 取得
    const data = mdl.getData("応募者", conditions);
    
    // ログ出力
    Logger.log(data);
    Logger.log("===============");
    Logger.log(data[0]["氏名"]);
    Logger.log(data[0]["年齢"]);
    Logger.log(data[0]["性別"]);
}
```

![get-sample-and.png](https://qiita-image-store.s3.ap-northeast-1.amazonaws.com/0/364477/fec4b48c-7cd7-cff8-275f-c91138c5dabb.png)

#### データのOR検索
もちろん、OR 条件でも絞り込みが可能です。
ただし、この場合は別のメソッド ```orGetData()``` を使用します。

```javascript:orGetData
orGetData(sheetName, conditions, offsetRow, limitRow)

[arg1] sheetName  : シート名
[arg2] conditions : 絞り込み条件
[arg3] offsetRow  : データ取得の開始数
[arg4] limitRow   : データ取得最大数
```

```javascript:データの OR 検索取得
function myFunction() {
    // インスタンスの生成
    const mdl = new Model();
    
    // 検索条件を設定
    const conditions = [
      { key: "氏名", value: "天野 結女" },
      { key: "氏名", value: "結城 莉央" },
      { key: "性別", value: "非公開" }
    ];
    
    // データの OR 取得
    const data = mdl.orGetData("応募者", conditions);
    
    // ログ出力
    Logger.log(data);
    Logger.log("===============");
    Logger.log(data[2]["氏名"]);
    Logger.log(data[2]["年齢"]);
    Logger.log(data[2]["性別"]);
}
```

![get-sample-or.png](https://qiita-image-store.s3.ap-northeast-1.amazonaws.com/0/364477/730cd192-7280-8ed4-e9c2-e43e8b8aaa33.png)

#### offset,limitでのデータ取得
```getData()```, ```orGetData()``` の両方のメソッドで offset, limit を設定できます。

```javascript:データの offset, limit 取得
function myFunction() {
    // インスタンスの生成
    const mdl = new Model();
    
    // offset, limit の設定（ここでは固定値）
    const offset = 2;
    const limit = 2;
    
    // 全データから offset, limit による取得
    const data = mdl.getData("応募者", null, offset, limit);
    
    // ログ出力
    Logger.log(data);
    Logger.log("===============");
    Logger.log(data[0]["氏名"]);
    Logger.log(data[0]["年齢"]);
    Logger.log(data[0]["性別"]);
}
```

![get-sample-offset-kimit.png](https://qiita-image-store.s3.ap-northeast-1.amazonaws.com/0/364477/5a411ed1-0d41-a493-ce38-f5095e28e388.png)

### データの新規登録
スプレッドシートの最終行に指定したデータの値を追加することができます。
```insertData()``` メソッドを使用します。

```javascript:insertData
insertData(sheetName, keyValuePairs)

[arg1] sheetName     : シート名
[arg2] keyValuePairs : 挿入するデータ (キーと値のペア)
```

```javascript:データの新規登録
function myFunction() {
    // インスタンスの生成
    const mdl = new Model();
    
    // 新規登録するデータ（キーと値のセット）リストを設定
    const keyValuePairs = [
      { "氏名": "村田 一生　", "年齢": "19", "性別": "男性" },
      { "氏名": "Ryan Hyeisang　", "年齢": "22", "性別": "女性" },
      { "氏名": "浜塚 悠　", "年齢": "14" } // 未設定項目は 空 で設定されます
    ]
    
    // データの新規登録
    const result = mdl.insertData("応募者", keyValuePairs);
    
    // ログ出力
    Logger.log(result); // 実行結果（Boolean）が返されます
}
```

![result-insert-sample.png](https://qiita-image-store.s3.ap-northeast-1.amazonaws.com/0/364477/7a2e011c-e575-7ebb-8174-ffbdd737a746.png)

![insert-spreadsheet-sample.png](https://qiita-image-store.s3.ap-northeast-1.amazonaws.com/0/364477/77c89952-f7c6-a400-20aa-b216a685ff14.png)

### データの更新
スプレッドシートに登録されているデータを検索し、対象データを更新することもできます。
データの検索は AND 条件で絞り込まれます。

```updateData()``` を使用します。

```javascript:updateData
updateData(sheetName, keyValuePairs, conditions)

[arg1] sheetName     : シート名
[arg2] keyValuePairs : 更新するデータ (キーと値のペア)
[arg3] conditions    : 絞り込み条件
```

![update-sample-before.png](https://qiita-image-store.s3.ap-northeast-1.amazonaws.com/0/364477/02f42c00-add5-b2e0-68a1-b89cf789ef0e.png)

```javascript:データの更新
function myFunction() {
    // インスタンスの生成
    const mdl = new Model();
    
    // 更新するデータ（キーと値のセット）を設定 : 必須
    const keyValuePair = { "性別": "女性" }; // 注意： オブジェクト、配列ではない

    // 更新するデータの絞り込み条件を設定 : 必須
    const conditions = [
      { key: "氏名", value: "浜塚 悠" },
      { key: "年齢", value: 14 }
    ];
    
    // データの更新
    const result = mdl.updateData("応募者", keyValuePair, conditions);
    
    // ログ出力
    Logger.log(result); // 実行結果（Boolean）が返されます
}
```

![update-data-sample-result.png](https://qiita-image-store.s3.ap-northeast-1.amazonaws.com/0/364477/fffcd0a7-a5e7-8d2a-cfd7-c48f1aa39452.png)

![update-sample-after.png](https://qiita-image-store.s3.ap-northeast-1.amazonaws.com/0/364477/ae224403-d12a-aede-cb93-070b7742e10a.png)

#### OR検索によるデータ更新
```orUpdateData()``` メソッドを用いると、絞り込み条件を OR としてデータを抽出し更新することもできます。

```javascript:orUpdateData
orUpdateData(sheetName, keyValuePairs, conditions)

[arg1] sheetName     : シート名
[arg2] keyValuePairs : 更新するデータ (キーと値のペア)
[arg3] conditions    : 絞り込み条件
```


![orUpdate-data-before.png](https://qiita-image-store.s3.ap-northeast-1.amazonaws.com/0/364477/d95d5ef2-e6f6-f567-d498-9b89da10a5b5.png)

```javascript:OR 検索によるデータ更新
function myFunction() {
    // インスタンスの生成
    const mdl = new Model();
    
    // 更新するデータ（キーと値のセット）を設定 : 必須
    const keyValuePair = { "年齢": "18" }; // 注意： オブジェクト、配列ではない

    // 更新するデータの絞り込み条件を設定 : 必須
    const conditions = [
      { key: "氏名", value: "天野 結女" },
      { key: "氏名", value: "結城 莉央" }
    ];
    
    // データの更新
    const result = mdl.orUpdateData("応募者", keyValuePair, conditions);
    
    // ログ出力
    Logger.log(result); // 実行結果（Boolean）が返されます
}
```

![orUpdate-data-after.png](https://qiita-image-store.s3.ap-northeast-1.amazonaws.com/0/364477/860f3ce6-171b-e2c8-71bb-21b1cec3783c.png)

### データの削除
スプレッドシートに登録されているデータを検索し、対象データを削除することもできます。
データの検索は AND 条件で絞り込まれます。

```deleteData()``` メソッドを使用します。

```javascript:deleteData
deleteData(sheetName, conditions)

[arg1] sheetName     : シート名
[arg2] conditions    : 絞り込み条件
```

![delete-sample-before.png](https://qiita-image-store.s3.ap-northeast-1.amazonaws.com/0/364477/cb8eebb3-097a-0031-7ac8-7bac4971feb4.png)

```javascript:データの削除
function myFunction() {
    // インスタンスの生成
    const mdl = new Model();

    // 更新するデータの絞り込み条件を設定 : 必須
    const conditions = [
      { key: "氏名", value: "浜塚 悠" },
      { key: "年齢", value: 14 }
    ];
    
    // データの削除
    const result = mdl.deleteData("応募者", conditions);
    
    // ログ出力
    Logger.log(result); // 実行結果（Boolean）が返されます
}
```

![delete-sample-after.png](https://qiita-image-store.s3.ap-northeast-1.amazonaws.com/0/364477/af1a708a-7927-f7d1-17e3-72fb9b3e6cb9.png)

#### OR検索によるデータの削除
```orDeleteData()``` メソッドを使用します。

```javascript:orDeleteData
orDeleteData(sheetName, conditions)

[arg1] sheetName     : シート名
[arg2] conditions    : 絞り込み条件
```

![orDelete-sample-before.png](https://qiita-image-store.s3.ap-northeast-1.amazonaws.com/0/364477/c02065bd-8b6b-7ef6-6fc7-36735e305d4c.png)

```javascript:OR 検索によるデータの削除
function myFunction() {
    // インスタンスの生成
    const mdl = new Model();

    // 更新するデータの絞り込み条件を設定 : 必須
    const conditions = [
      { key: "年齢", value: 18 },
      { key: "年齢", value: 22 }
    ];
    
    // データの削除
    const result = mdl.orDeleteData("応募者", conditions);
    
    // ログ出力
    Logger.log(result); // 実行結果（Boolean）が返されます
}
```

![orDelete-sample-after.png](https://qiita-image-store.s3.ap-northeast-1.amazonaws.com/0/364477/c0e62e20-b3a6-e0a8-d161-79e633b05773.png)

### データの全削除
ヘッダー以外の全てのデータ行を削除します。
```truncateData()``` メソッドを使用します。

```javascript:truncateData
truncateData(sheetName)

[arg1] sheetName     : シート名
```

```javascript:データの全削除（ヘッダー以外）
function myFunction() {
    // インスタンスの生成
    const mdl = new Model();

    // データの全削除（ヘッダーを除く）
    const result = mdl.truncateData("応募者");
    
    // ログ出力
    Logger.log(result); // 実行結果（Boolean）が返されます
}
```

![truncate-sample.png](https://qiita-image-store.s3.ap-northeast-1.amazonaws.com/0/364477/da3ebf3b-310f-3b16-4649-82ac20523f35.png)

### 比較演算子によるデータの検索
最後に、少し応用編となりますが、データの絞り込みを比較演算子で行うこともできます。
方法は条件を設定するオブジェクト配列の key に比較演算子を含めて設定するものとなります。

比較演算子は、```>```, ```>=```, ```<```, ```<=```, ```!=```, ```!==```, ```==``` を含めることができます。

:::note info
比較演算子を指定しない通常の絞り込みは ```===``` がデフォルトとなっています。
:::

:::note info
この比較演算子によるデータの絞り込みは、前述したデータの更新・削除においても適用できます。
:::

```javascript:比較演算子によるデータ取得
function myFunction() {
    // インスタンスの生成
    const mdl = new Model();
    
    // 検索条件を設定
    const conditions = [
      { key: "年齢 >", value: 30 }, // 比較演算子の前に半角スペースを入れる
      { key: "年齢 <=", value: 18 } // 比較演算子の前に半角スペースを入れる
    ];
    
    // データの OR 取得
    const data = mdl.orGetData("応募者", conditions);
    
    // ログ出力
    Logger.log(data);
    Logger.log("===============");
    Logger.log(data[0]["氏名"]);
    Logger.log(data[0]["年齢"]);
    Logger.log(data[0]["性別"]);
}
```
![get-operator-sample.png](https://qiita-image-store.s3.ap-northeast-1.amazonaws.com/0/364477/767c05b4-7f4b-7b13-6b33-f6fc16289e59.png)

# おわりに
自分でも使用する中で機能の修正や追加をしてブラッシュアップしていきたいと思います。
最後までお読みいただきありがとうございました。
