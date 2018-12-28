---
title: Office ライブラリの JavaScript API をそのコンテンツ配信ネットワーク (CDN) から参照する
description: ''
ms.date: 01/23/2018
ms.openlocfilehash: 9943af86419652e5f5e89b1741b32b4e0da15e77
ms.sourcegitcommit: 60fd8a3ac4a6d66cb9e075ce7e0cde3c888a5fe9
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/28/2018
ms.locfileid: "27457622"
---
# <a name="referencing-the-javascript-api-for-office-library-from-its-content-delivery-network-cdn"></a><span data-ttu-id="98753-102">Office ライブラリの JavaScript API をそのコンテンツ配信ネットワーク (CDN) から参照する</span><span class="sxs-lookup"><span data-stu-id="98753-102">Referencing the JavaScript API for Office library from its content delivery network (CDN)</span></span>

> [!NOTE]
> <span data-ttu-id="98753-103">この記事で説明している手順に加え、TypeScript を使用する場合には、ノードが有効になっているシステム プロンプト (または git bash ウィンドウ) でプロジェクト フォルダーのルートから次のコマンドを実行して、Intellisense を入手する必要があります。</span><span class="sxs-lookup"><span data-stu-id="98753-103">In addition to the steps described in this article, if you want to use TypeScript, then to get Intellisense you will need run the following command in a Node-enabled system prompt (or git bash window) from the root of your project folder.</span></span> <span data-ttu-id="98753-104">(npm を含む) [Node.js](https://nodejs.org) をインストールしておく必要があります。</span><span class="sxs-lookup"><span data-stu-id="98753-104">You must have [Node.js](https://nodejs.org) installed (which includes npm).</span></span>
> 
> ```bash
> npm install --save-dev @types/office-js
> ```

<span data-ttu-id="98753-105">[JavaScript API for Office](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office) ライブラリは、Office.js ファイルと関連するホスト アプリケーション固有の .js ファイル (Excel-15.js や Outlook-15.js など) で構成されています。</span><span class="sxs-lookup"><span data-stu-id="98753-105">The [JavaScript API for Office](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office) library consists of the Office.js file and associated host application-specific .js files, such as Excel-15.js and Outlook-15.js.</span></span> 


<span data-ttu-id="98753-106">最も簡単に API を参照する方法は、次に示す `<script>` をページの `<head>` タグに追加して、CDN を使用することです。</span><span class="sxs-lookup"><span data-stu-id="98753-106">The simplest way to reference the API is to use our CDN by adding the following `<script>` to your page's `<head>` tag:</span></span>  

```html
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
```

<span data-ttu-id="98753-p102">CDN URL で `office.js` の前にある `/1/` は、Office.js のバージョン 1 内で最新の増分リリースを指定します。JavaScript API for Office が旧バージョンとの互換性を維持するので、最新リリースはバージョン 1 で以前導入されていた API メンバーを引き続きサポートします。既存のプロジェクトをアップグレードする必要がある場合は、「[JavaScript API for Office およびマニフェスト スキーマ ファイルのバージョンを更新する](update-your-javascript-api-for-office-and-manifest-schema-version.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="98753-p102">The  `/1/` in front of `office.js` in the CDN URL specifies the latest incremental release within version 1 of Office.js. Because the JavaScript API for Office maintains backward compatibility, the latest release will continue to support API members that were introduced earlier in version 1. If you need to upgrade an existing project, see [Update the version of your JavaScript API for Office and manifest schema files](update-your-javascript-api-for-office-and-manifest-schema-version.md).</span></span> 

<span data-ttu-id="98753-p103">AppSource から Office アドインを発行する場合は、この CDN の参照を使用する必要があります。ローカル参照は、内部シナリオ、開発シナリオ、デバッグ シナリオにのみ適用できます。</span><span class="sxs-lookup"><span data-stu-id="98753-p103">If you plan to publish your Office Add-in from AppSource, you must use this CDN reference. Local references are only appropriate for internal, development, and debugging scenarios.</span></span>

> [!IMPORTANT]
>  <span data-ttu-id="98753-p104">Office ホスト アプリケーションのアドインを開発する場合は、ページの `<head>` セクションの内側から JavaScript API for Office を参照します。これにより、あらゆる body 要素の前に API が完全に初期化されます。Office ホストでは、アクティブ化の 5 秒以内にアドインを初期化する必要があります。このしきい値内にアドインがアクティブにならない場合は、応答なしが宣言され、エラー メッセージがユーザーに表示されます。</span><span class="sxs-lookup"><span data-stu-id="98753-p104">When you develop an add-in for any Office host application, reference the JavaScript API for Office from inside the `<head>` section of the page. This ensures that the API is fully initialized prior to any body elements. Office hosts require that add-ins initialize within 5 seconds of activation. If your add-in doesn't activate within this threshold, it will be declared unresponsive and an error message will be displayed to the user.</span></span>       

## <a name="see-also"></a><span data-ttu-id="98753-116">関連項目</span><span class="sxs-lookup"><span data-stu-id="98753-116">See also</span></span>

- [<span data-ttu-id="98753-117">JavaScript API for Office について</span><span class="sxs-lookup"><span data-stu-id="98753-117">Understanding the JavaScript API for Office</span></span>](understanding-the-javascript-api-for-office.md)    
- [<span data-ttu-id="98753-118">JavaScript API for Office</span><span class="sxs-lookup"><span data-stu-id="98753-118">JavaScript API for Office</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office)
    
