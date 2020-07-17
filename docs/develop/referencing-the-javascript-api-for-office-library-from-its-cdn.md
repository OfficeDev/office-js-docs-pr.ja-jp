---
title: Office JavaScript API ライブラリの参照
description: アドインで Office JavaScript API ライブラリおよび型定義を参照する方法について説明します。
ms.date: 06/23/2020
localization_priority: Normal
ms.openlocfilehash: 3f90b0798b14b66fe6d01f62eca3802fce179bec
ms.sourcegitcommit: a4873c3525c7d30ef551545d27eb2c0a16b4eb50
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/25/2020
ms.locfileid: "44888132"
---
# <a name="referencing-the-office-javascript-api-library"></a><span data-ttu-id="d36c2-103">Office JavaScript API ライブラリの参照</span><span class="sxs-lookup"><span data-stu-id="d36c2-103">Referencing the Office JavaScript API library</span></span>

<span data-ttu-id="d36c2-104">[Office JAVASCRIPT API](../reference/javascript-api-for-office.md)ライブラリには、アドインが office ホストと対話するために使用できる api が用意されています。</span><span class="sxs-lookup"><span data-stu-id="d36c2-104">The [Office JavaScript API](../reference/javascript-api-for-office.md) library provides the APIs that your add-in can use to interact with the Office host.</span></span> <span data-ttu-id="d36c2-105">ライブラリを参照する最も簡単な方法は、 `<script>` `<head>` HTML ページのセクション内に次のタグを追加することによって、コンテンツ配信ネットワーク (CDN) を使用する方法です。</span><span class="sxs-lookup"><span data-stu-id="d36c2-105">The simplest way to reference the library is to use the content delivery network (CDN) by adding the following `<script>` tag within the `<head>` section of your HTML page:</span></span>  

```html
<head>
    ...
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
</head>
```

<span data-ttu-id="d36c2-106">これにより、アドインが最初に読み込まれたときに Office JavaScript API ファイルがダウンロードされてキャッシュされるので、指定されたバージョンの Office.js と関連付けられたファイルの最新の実装が使用されていることを確認してください。</span><span class="sxs-lookup"><span data-stu-id="d36c2-106">This will download and cache the Office JavaScript API files the first time your add-in loads to make sure that it is using the most up-to-date implementation of Office.js and its associated files for the specified version.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="d36c2-107">ページのセクションの内側から Office JavaScript API を参照して、 `<head>` API が body 要素の前に完全に初期化されていることを確認する必要があります。</span><span class="sxs-lookup"><span data-stu-id="d36c2-107">You must reference the Office JavaScript API from inside the `<head>` section of the page to ensure that the API is fully initialized prior to any body elements.</span></span> <span data-ttu-id="d36c2-108">Office ホストでは、アクティブ化の 5 秒以内にアドインを初期化する必要があります。</span><span class="sxs-lookup"><span data-stu-id="d36c2-108">Office hosts require that add-ins initialize within 5 seconds of activation.</span></span> <span data-ttu-id="d36c2-109">このしきい値内にアドインがアクティブにならない場合は、応答なしが宣言され、エラー メッセージがユーザーに表示されます。</span><span class="sxs-lookup"><span data-stu-id="d36c2-109">If your add-in doesn't activate within this threshold, it will be declared unresponsive and an error message will be displayed to the user.</span></span>

## <a name="api-versioning-and-backward-compatibility"></a><span data-ttu-id="d36c2-110">API のバージョン管理と下位互換性</span><span class="sxs-lookup"><span data-stu-id="d36c2-110">API versioning and backward compatibility</span></span>

<span data-ttu-id="d36c2-111">前の HTML スニペットで、 `/1/` CDN URL の前の部分には `office.js` Office.js のバージョン1で最新の増分リリースが指定されています。</span><span class="sxs-lookup"><span data-stu-id="d36c2-111">In the previous HTML snippet, the `/1/` in front of `office.js` in the CDN URL specifies the latest incremental release within version 1 of Office.js.</span></span> <span data-ttu-id="d36c2-112">Office JavaScript API は下位互換性を維持しているため、最新のリリースでは、以前のバージョン1で導入された API メンバーを引き続きサポートしています。</span><span class="sxs-lookup"><span data-stu-id="d36c2-112">Because the Office JavaScript API maintains backward compatibility, the latest release will continue to support API members that were introduced earlier in version 1.</span></span> <span data-ttu-id="d36c2-113">既存のプロジェクトをアップグレードする必要がある場合は、「 [Office JAVASCRIPT API およびマニフェストスキーマファイルのバージョンを更新](update-your-javascript-api-for-office-and-manifest-schema-version.md)する」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="d36c2-113">If you need to upgrade an existing project, see [Update the version of your Office JavaScript API and manifest schema files](update-your-javascript-api-for-office-and-manifest-schema-version.md).</span></span> 

<span data-ttu-id="d36c2-p104">AppSource から Office アドインを発行する場合は、この CDN の参照を使用する必要があります。ローカル参照は、内部シナリオ、開発シナリオ、デバッグ シナリオにのみ適用できます。</span><span class="sxs-lookup"><span data-stu-id="d36c2-p104">If you plan to publish your Office Add-in from AppSource, you must use this CDN reference. Local references are only appropriate for internal, development, and debugging scenarios.</span></span>

> [!NOTE]
> <span data-ttu-id="d36c2-116">プレビュー API を使用するには、CDN (`https://appsforoffice.microsoft.com/lib/beta/hosted/office.js`) にある Office JavaScript API ライブラリのプレビュー バージョンを参照します。</span><span class="sxs-lookup"><span data-stu-id="d36c2-116">To use preview APIs, reference the preview version of the Office JavaScript API library on the CDN: `https://appsforoffice.microsoft.com/lib/beta/hosted/office.js`.</span></span>

## <a name="enabling-intellisense-for-a-typescript-project"></a><span data-ttu-id="d36c2-117">TypeScript プロジェクトに対して IntelliSense を有効にする</span><span class="sxs-lookup"><span data-stu-id="d36c2-117">Enabling IntelliSense for a TypeScript project</span></span>

<span data-ttu-id="d36c2-118">前述したように Office JavaScript API を参照するだけでなく、[指定](https://github.com/DefinitelyTyped/DefinitelyTyped/tree/master/types/office-js)した型定義を使用して TypeScript アドインプロジェクトの IntelliSense を有効にすることもできます。</span><span class="sxs-lookup"><span data-stu-id="d36c2-118">In addition to referencing the Office JavaScript API as described previously, you can also enable IntelliSense for TypeScript add-in project by using the type definitions from [DefinitelyTyped](https://github.com/DefinitelyTyped/DefinitelyTyped/tree/master/types/office-js).</span></span> <span data-ttu-id="d36c2-119">これを行うには、プロジェクトフォルダーのルートから、ノードが有効なシステムプロンプト (または git bash ウィンドウ) で次のコマンドを実行します。</span><span class="sxs-lookup"><span data-stu-id="d36c2-119">To do so, run the following command in a Node-enabled system prompt (or git bash window) from the root of your project folder.</span></span> <span data-ttu-id="d36c2-120">(npm を含む) [Node.js](https://nodejs.org) をインストールしておく必要があります。</span><span class="sxs-lookup"><span data-stu-id="d36c2-120">You must have [Node.js](https://nodejs.org) installed (which includes npm).</span></span>

```command&nbsp;line
npm install --save-dev @types/office-js
```

## <a name="preview-apis"></a><span data-ttu-id="d36c2-121">プレビュー Api</span><span class="sxs-lookup"><span data-stu-id="d36c2-121">Preview APIs</span></span>

<span data-ttu-id="d36c2-122">新しい JavaScript Api が最初に "プレビュー" で導入され、さらにテストが行われ、ユーザーフィードバックが必要になった後、特定の番号付き要件セットの一部となります。</span><span class="sxs-lookup"><span data-stu-id="d36c2-122">New JavaScript APIs are first introduced in "preview" and later become part of a specific numbered requirement set after sufficient testing occurs and user feedback is required.</span></span>

[!INCLUDE [Information about using preview APIs](../includes/using-preview-apis-host.md)]

## <a name="see-also"></a><span data-ttu-id="d36c2-123">関連項目</span><span class="sxs-lookup"><span data-stu-id="d36c2-123">See also</span></span>

- [<span data-ttu-id="d36c2-124">Office JavaScript API について</span><span class="sxs-lookup"><span data-stu-id="d36c2-124">Understanding the Office JavaScript API</span></span>](understanding-the-javascript-api-for-office.md)
- [<span data-ttu-id="d36c2-125">Office の JavaScript API</span><span class="sxs-lookup"><span data-stu-id="d36c2-125">Office JavaScript API</span></span>](../reference/javascript-api-for-office.md)
