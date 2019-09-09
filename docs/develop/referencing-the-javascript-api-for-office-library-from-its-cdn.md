---
title: Office ライブラリの JavaScript API をそのコンテンツ配信ネットワーク (CDN) から参照する
description: ''
ms.date: 03/19/2019
localization_priority: Priority
ms.openlocfilehash: 6945cb9e2e93209c1568575d8c393cf00ae47431
ms.sourcegitcommit: d34aa0b282cc76ffff579da2a7945efd12fb7340
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/05/2019
ms.locfileid: "36769583"
---
# <a name="referencing-the-javascript-api-for-office-library-from-its-content-delivery-network-cdn"></a>Office ライブラリの JavaScript API をそのコンテンツ配信ネットワーク (CDN) から参照する

> [!NOTE]
> この記事で説明している手順に加え、TypeScript を使用する場合には、ノードが有効になっているシステム プロンプト (または git bash ウィンドウ) でプロジェクト フォルダーのルートから次のコマンドを実行して、IntelliSense を入手する必要があります。 (npm を含む) [Node.js](https://nodejs.org) をインストールしておく必要があります。
> 
> ```command&nbsp;line
> npm install --save-dev @types/office-js
> ```

[JavaScript API for Office](/office/dev/add-ins/reference/javascript-api-for-office) ライブラリは、Office.js ファイルと関連するホスト アプリケーション固有の .js ファイル (Excel-15.js や Outlook-15.js など) で構成されています。 


最も簡単に API を参照する方法は、次に示す `<script>` をページの `<head>` タグに追加して、CDN を使用することです。  

```html
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
```

CDN URL で `office.js` の前にある `/1/` は、Office.js のバージョン 1 内で最新の増分リリースを指定します。JavaScript API for Office が旧バージョンとの互換性を維持するので、最新リリースはバージョン 1 で以前導入されていた API メンバーを引き続きサポートします。既存のプロジェクトをアップグレードする必要がある場合は、「[JavaScript API for Office およびマニフェスト スキーマ ファイルのバージョンを更新する](update-your-javascript-api-for-office-and-manifest-schema-version.md)」を参照してください。 

AppSource から Office アドインを発行する場合は、この CDN の参照を使用する必要があります。ローカル参照は、内部シナリオ、開発シナリオ、デバッグ シナリオにのみ適用できます。

> [!IMPORTANT]
> Office ホスト アプリケーションのアドインを開発する場合は、ページの `<head>` セクションの内側から JavaScript API for Office を参照します。これにより、あらゆる body 要素の前に API が完全に初期化されます。Office ホストでは、アクティブ化の 5 秒以内にアドインを初期化する必要があります。このしきい値内にアドインがアクティブにならない場合は、応答なしが宣言され、エラー メッセージがユーザーに表示されます。

## <a name="see-also"></a>関連項目

- [JavaScript API for Office について](understanding-the-javascript-api-for-office.md)
- [JavaScript API for Office](/office/dev/add-ins/reference/javascript-api-for-office)
