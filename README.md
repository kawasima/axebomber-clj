Axebomber
=========

Excel方眼紙を出力するためのライブラリです。

## Excel方眼紙の定義

基本は列幅をベースフォントの2文字分、行幅を1文字分に固定してレイアウトされたExcelシートを指します。
細かい点で宗派がわかれるのでAxebomberが扱う方眼紙を定義しておきます。

* セル内で改行はしない。改行したいときは次の行に書き込む。
* 1マス1文字はフォントによって文字切れを起こしやすいので対応しない。
* 数値などセルを右寄せしたい場合には、セル結合をする。それ以外はセル結合は使わない。

Excelがドキュメント作成に好んで使われるのは、広大なキャンバスに適当に
レイアウティングした文書を作っても、印刷範囲の設定でA4用紙一杯にキレイに
印刷されることが大きな理由と考えます。

この"適当なレイアウト"の部分を分解すると、以下のような性質がみられます。

* 基本的には上から下に文書の構成要素を配置していく、垂直方向のフローティングである。
* A4横一枚に収めるためには、上記垂直フローティングで、段組みを構成することもある。
*

したがってAxebomberもこのデザインに沿った、レイアウトが出来るよう設計されています。

## Axebomberの基本操作

### 決まった位置に書き込む

```clojure
(render sheet 2 2 "こんにちは、Excel")
```

このようにrender関数を使って指定したマス目に書き込むことができます。

### 表組み

HTMLと同じように、:table, :tr, :th, :td といったタグを使って表現します。
属性に:sizeを与えると、その分だけ横幅が確保されます。
行の高さは、その行に含まれるセルの高さの最大値に自動的にセットされるので、特に指定する必要はありません。

```clojure
(render sheet 2 2
  [:table
    [:tr
      [:th {:size 3} "ID"]
      [:th {:size 8} "名称"]]
    [:tr
      [:td "1"]
      [:td "kawasima"]]])
```

#### センタリング／右寄せ

デフォルトはセル内のテキストは左寄せになりますが、:tdタグの属性に`{:text-align "center"}`または`{:text-align "right"}`を指定すると、 セルが結合されて中央寄せ、右寄せされます。

```clojure
(render sheet 2 2
  [:table
    [:tr
      [:th {:size 3} "ID"]
      [:th {:size 8} "金額"]]
    [:tr
      [:td "1"]
      [:td {:text-align "right"} 15000]]])
```

#### 複数のセルを結合する

HTMLのテーブルと同じようにcolspan/rowspan属性が利用できます。

```clojure
(render sheet 2 2
  [:table
    [:tr
      [:th {:size 3} "大分類"]
      [:th {:size 8} "小分類"]]
    [:tr
      [:td {:colspan "2"} "マージ"]]])
```

### 箇条書き

HTMLと同じ要領で、:ul, :ol, :li のタグを使って表現します。
行頭文字に1マス使われます。

```clojure
(render sheet 2 2
  [:ul {:list-style "・"}
    [:li "りんご"]
    [:li "みかん"]
    [:li "ばなな"]])
```

## プログラマブルな使い方

これらのAPIはClojureのHTMLテンプレートとして使われているhiccupをベースにしています。そしてhiccupと同じように、forやifの構文が使えます。
```clojure
(render sheet 2 2
  [:table
    (for [item items]
      [:tr
        [:td (:id item)]
        [:td (:name item)]])])
```

## ライセンス

Source Copyright © 2014 kawasima.
Distributed under the Eclipse Public License, the same as Clojure uses.
