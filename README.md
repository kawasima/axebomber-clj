Axebomber
=========

Excel方眼紙を出力するためのライブラリです。

## Excel方眼紙の定義

基本は列幅をベースフォントの2文字分、行幅を1文字分に固定してレイアウトされたExcelシートを指します。
細かい点で宗派がわかれるのでAxebomberが扱う方眼紙を定義しておきます。

* セル内で改行はしない。改行したいときは次の行に書き込む。
* 1マス1文字はフォントによって文字切れを起こしやすいので対応しない。
* 数値などセルを右寄せしたい場合には、セル結合をする。
  それ以外はセル結合は使わない。

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
      [:td "名称"]
      [:td "名称"]]
      ])
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

## 方眼紙レイアウトの注意点

ブロックが横にフローティングされるHTMLと異なり、デフォルトは縦にフローティングされます。



## ライセンス

Source Copyright © 2014 kawasima.
Distributed under the Eclipse Public License, the same as Clojure uses.
