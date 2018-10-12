---
title: Office ライブラリの JavaScript API を Office コンテンツ配信ネットワーク (CDN) から参照する
description: ''
ms.date: 01/23/2018
ms.openlocfilehash: 0ad589ee98342ee72259cddc0957277e9018f186
ms.sourcegitcommit: c53f05bbd4abdfe1ee2e42fdd4f82b318b363ad7
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/12/2018
ms.locfileid: "25505420"
---
# <a name="referencing-the-javascript-api-for-office-library-from-its-content-delivery-network-cdn"></a>Office ライブラリの JavaScript API を Office コンテンツ配信ネットワーク (CDN) から参照する

> [!NOTE]
> この記事で説明する手順に加えて TypeScript を使用する場合は、 Intellisense を取得するために、プロジェクト フォルダーのルートからノード対応のシステム プロンプト (または git bash ウィンドウ) で次のコマンドを実行する必要があります。[Node.js](https://nodejs.org) がインストールされている必要があります (npm を含む)。
> 
> ```
> npm install --save-dev @types/office-js
> ```

[JavaScript API for Office](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office?view=office-js) ライブラリは、Office.js ファイルと関連するホスト アプリケーション固有の .js ファイル (Excel-15.js や Outlook-15.js など) で構成されています。 


最も簡単に API を参照する方法は、次に示す `<script>` をページの `<head>` タグに追加して、CDN を使用することです。  

```html
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
```

CDN URL で `office.js` の前にある `/1/` は、Office.js のバージョン 1 内で最新の増分リリースを指定します。JavaScript API for Office は旧バージョンとの互換性を維持しているため、最新リリースにおいてもバージョン 1 の時点で導入されていた API メンバーを引き続きサポートします。既存のプロジェクトをアップグレードする必要がある場合は、「[JavaScript API for Office とマニフェスト スキーマ ファイルのバージョンを更新する](update-your-javascript-api-for-office-and-manifest-schema-version.md)」を参照してください。 

AppSource から Office アドインを発行する場合は、この CDN の参照を使用する必要があります。ローカル参照は、内部シナリオ、開発シナリオ、デバッグ シナリオにのみ適用できます。

> [!IMPORTANT]
>  Office ホスト アプリケーションのアドインを開発する場合は、ページの `<head>` セクションの中から JavaScript API for Office を参照します。これにより、すべての body 要素の前に API が完全に初期化されます。Office ホストでは、アクティブ化から 5 秒以内にアドインを初期化する必要があります。このしきい値内にアドインがアクティブにならない場合は、応答なしが宣言され、エラー メッセージがユーザーに表示されます。       

## <a name="see-also"></a>関連項目

- [JavaScript API for Office について](understanding-the-javascript-api-for-office.md)    
- [JavaScript API for Office](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office?view=office-js)
    
