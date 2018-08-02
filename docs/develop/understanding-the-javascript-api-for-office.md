---
title: JavaScript API for Office について
description: ''
ms.date: 01/23/2018
ms.openlocfilehash: fef2cdad69408f099296461066f1ea380e3b118b
ms.sourcegitcommit: bc68b4cf811b45e8b8d1cbd7c8d2867359ab671b
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/02/2018
ms.locfileid: "21703813"
---
# <a name="understanding-the-javascript-api-for-office"></a>JavaScript API for Office について

この記事では、JavaScript API for Office とその使用方法に関する情報を提供します。参照情報については、「[JavaScript API for Office](https://dev.office.com/reference/add-ins/javascript-api-for-office)」を参照してください。Visual Studio プロジェクト ファイルを JavaScript API for Office の最新バージョンに更新する方法については、「[JavaScript API for Office およびマニフェスト スキーマ ファイルのバージョンを更新する](update-your-javascript-api-for-office-and-manifest-schema-version.md)」を参照してください。

> [!NOTE]
> AppSource にアドインを[公開](../publish/publish.md)し、Office エクスペリエンスで利用できるようにする予定がある場合は、[AppSource の検証ポリシー](https://docs.microsoft.com/en-us/office/dev/store/validation-policies)に準拠していることを確認してください。たとえば、検証に合格するには、定義したメソッドをサポートするすべてのプラットフォームでアドインが動作する必要があります (詳細については、[セクション 4.12](https://docs.microsoft.com/en-us/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) と [Office アドインを使用できるホストおよびプラットフォーム](../overview/office-add-in-availability.md)のページを参照してください)。 

## <a name="referencing-the-javascript-api-for-office-library-in-your-add-in"></a>アドインで JavaScript API for Office ライブラリを参照する

[JavaScript API for Office](https://dev.office.com/reference/add-ins/javascript-api-for-office) ライブラリは、Office.js ファイルと関連するホスト アプリケーション固有のファイル (Excel-15.js や Outlook-15.js など) で構成されています。最も簡単に API を参照する方法は、次に示す `<script>` をページの `<head>` タグに追加して、CDN を使用することです。  

```html
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
```

これにより、アドインが最初に読み込まれるときに JavaScript API for Office ファイルのダウンロードとキャッシュを実行して、アドインが確実に指定したバージョンの最新の Office.js および関連ファイルを使用するようにします。

バージョン管理や下位互換性の処理方法など、Office.js CDN に関する詳細については、「[ JavaScript API for Office ライブラリをそのコンテンツ配信ネットワーク (CDN) から参照する](referencing-the-javascript-api-for-office-library-from-its-cdn.md)」を参照してください。

## <a name="initializing-your-add-in"></a>アドインの初期化

**適用対象:** すべてのアドインの種類

Office.js は、API が完全に読み込まれていてユーザーによる操作ができる状態になっているときに起動されたとしても初期化を提供します。**initialize** イベント ハンドラーを使用すると、ユーザーに Excel のセルを複数選択するように求めるメッセージを表示し、選択された値で初期化したグラフを挿入するなど、アドインの一般的な初期化シナリオを実装できるようになります。また、アドインのその他のカスタム ロジックを初期化する場合 (バインドを確立する場合やアドインの既定の設定値を入力するように求めるメッセージを表示する場合) にも、initialize イベント ハンドラーが使用できます。

最小限の initialize イベントは、次の例のようになります。     

```js
Office.initialize = function () { };
```
独自の初期化ハンドラーやテストを含む追加の JavaScript フレームワークを使用している場合、そのようなフレームワークは Office.initialize イベントの内側に配置する必要があります。たとえば、[JQuery](https://jquery.com) の `$(document).ready()` 関数は次のように参照します。

```js
Office.initialize = function () {
    // Office is ready
    $(document).ready(function () {        
        // The document is ready
    });
  };
```

Office アドイン内のすべてのページで、初期化イベント **Office.initialize** にイベント ハンドラーを割り当てる必要があります。イベント ハンドラーを割り当てないと、アドインの起動時にエラーが発生することがあります。また、ユーザーが Excel Online、PowerPoint Online、Outlook Web App などの Office Online Web クライアントでアドインを使用しようとすると、アドインの実行が失敗します。初期化コードが必要ない場合は、上の最初の例のように、**Office.initialize** に割り当てる関数の本体を空にできます。

アドインの初期化時のイベントのシーケンスの詳細については、「[DOM とランタイム環境を読み込む](loading-the-dom-and-runtime-environment.md)」を参照してください。

#### <a name="initialization-reason"></a>初期化の理由
作業ウィンドウ アドインとコンテンツ アドインについては、Office.initialize に追加の _reason_ パラメーターを使用できます。このパラメーターは、アドインがどのように現在のドキュメントに追加されたかを判断するために使用できます。これは、最初にアドインが挿入されたときと、既にアドインがドキュメント内に存在しているときに別のロジックを提供するために使用できます。 

```js
Office.initialize = function (reason) {
    $(document).ready(function () {
        switch (reason) {
            case 'inserted': console.log('The add-in was just inserted.');
            case 'documentOpened': console.log('The add-in is already part of the document.');
        }
    });
 };
```
詳細については、「[Office.initialize イベント](https://dev.office.com/reference/add-ins/shared/office.initialize)」および「[InitializationReason 列挙型](https://dev.office.com/reference/add-ins/shared/initializationreason-enumeration)」を参照してください。 

## <a name="office-javascript-api-object-model"></a>Office JavaScript オブジェクト モデル

初期化されると、アドインはホスト (たとえば Excel、Outlook など) を操作できるようになります。 特定の使用パターンに関する詳細については、「[Office JavaScript API オブジェクトモデル](/office-javascript-api-object-model.md)」ページを参照してください。 [共有 API](https://dev.office.com/reference/add-ins/javascript-api-for-office) および特定のホスト両方についても、詳細な参照ドキュメントがあります。

## <a name="api-support-matrix"></a>API サポート マトリックス


次の表は、アドインの種類 (コンテンツ、作業ウィンドウ、および Outlook) 全体でサポートされている API と機能、および [1.1 アドイン マニフェスト スキーマと機能 (JavaScript API for Office v1.1 でサポート)](update-your-javascript-api-for-office-and-manifest-schema-version.md) を使用してアドインがサポートする Office のホスト アプリケーションを指定する際に、これらの API と機能をホストする Office アプリケーションについてまとめたものです。


|||||||||
|:-----|:-----|:-----:|:-----:|:-----:|:-----:|:-----:|:-----:|
||**ホスト名**|データベース|ブック|メールボックス|プレゼンテーション|ドキュメント|プロジェクト|
||**サポートされる****ホスト アプリケーション**|Access Web アプリ|Excel、<br/>Excel Online|Outlook、<br/>Outlook Web App、<br/>デバイス用OWA|PowerPoint、<br/>PowerPoint Online|Word|プロジェクト|
|**サポートされるアドインの種類**|コンテンツ|Y|Y||Y|||
||作業ウィンドウ||Y||Y|Y|Y|
||Outlook|||Y||||
|**サポートされている API 機能**|テキストの読み取り/書き込み||Y||Y|Y|Y<br/>(読み取り専用)|
||マトリックスの読み取り/書き込み||Y|||Y||
||テーブルの読み取り/書き込み||Y|||Y||
||HTML の読み取り/書き込み|||||Y||
||読み取り/書き込み<br/>Office Open XML|||||Y||
||タスク、リソース、ビュー、フィールド プロパティの読み取り||||||Y|
||選択変更イベント||Y|||Y||
||ドキュメント全体の取得||||Y|Y||
||バインドとイベント バインド|Y<br/>(完全および部分的なテーブル バインドのみ)|Y|||Y||
||カスタム XML パーツの読み取り/書き込み|||||Y||
||アドイン状態データの保持 (設定)|Y<br/>(ホスト アドインごと)|Y<br/>(ドキュメントごと)|Y<br/>(メールボックスごと)|Y<br/>(ドキュメントごと)|Y<br/>(ドキュメントごと)||
||設定変更イベント|Y|Y||Y|Y||
||アクティブ ビュー モード<br/>およびビュー変更イベントの取得||||Y|||
||ドキュメント内の<br/>場所に移動||Y||Y|Y||
||ルールと RegEx を使用した<br/>文脈からのアクティブ化|||Y||||
||アイテム プロパティの読み取り|||Y||||
||ユーザー プロファイルの読み取り|||Y||||
||添付ファイルの取得|||Y||||
||ユーザー ID トークンの取得|||Y||||
||Exchange Web サービスの呼出|||Y||||
