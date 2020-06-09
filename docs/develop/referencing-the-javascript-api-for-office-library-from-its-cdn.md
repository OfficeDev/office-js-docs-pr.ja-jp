---
title: Office JavaScript API ライブラリの参照
description: アドインで Office JavaScript API ライブラリおよび型定義を参照する方法について説明します。
ms.date: 02/27/2020
localization_priority: Normal
ms.openlocfilehash: 8bd011c140ce61581ad4b1d06a43b04ad437f5c7
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/08/2020
ms.locfileid: "44609388"
---
# <a name="referencing-the-office-javascript-api-library"></a><span data-ttu-id="e72f9-103">Office JavaScript API ライブラリの参照</span><span class="sxs-lookup"><span data-stu-id="e72f9-103">Referencing the Office JavaScript API library</span></span>

<span data-ttu-id="e72f9-104">[Office JAVASCRIPT API](../reference/javascript-api-for-office.md)ライブラリには、アドインが office ホストと対話するために使用できる api が用意されています。</span><span class="sxs-lookup"><span data-stu-id="e72f9-104">The [Office JavaScript API](../reference/javascript-api-for-office.md) library provides the APIs that your add-in can use to interact with the Office host.</span></span> <span data-ttu-id="e72f9-105">ライブラリを参照する最も簡単な方法は、 `<script>` `<head>` HTML ページのセクション内に次のタグを追加することによって、コンテンツ配信ネットワーク (CDN) を使用する方法です。</span><span class="sxs-lookup"><span data-stu-id="e72f9-105">The simplest way to reference the library is to use the content delivery network (CDN) by adding the following `<script>` tag within the `<head>` section of your HTML page:</span></span>  

```html
<head>
    ...
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
</head>
```

<span data-ttu-id="e72f9-106">これにより、アドインが最初に読み込まれたときに Office JavaScript API ファイルがダウンロードされてキャッシュされるため、指定されたバージョンの Office .js および関連付けられたファイルの最新の実装を使用していることを確認できます。</span><span class="sxs-lookup"><span data-stu-id="e72f9-106">This will download and cache the Office JavaScript API files the first time your add-in loads to make sure that it is using the most up-to-date implementation of Office.js and its associated files for the specified version.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="e72f9-107">ページのセクションの内側から Office JavaScript API を参照して、 `<head>` API が body 要素の前に完全に初期化されていることを確認する必要があります。</span><span class="sxs-lookup"><span data-stu-id="e72f9-107">You must reference the Office JavaScript API from inside the `<head>` section of the page to ensure that the API is fully initialized prior to any body elements.</span></span> <span data-ttu-id="e72f9-108">Office ホストでは、アクティブ化の 5 秒以内にアドインを初期化する必要があります。</span><span class="sxs-lookup"><span data-stu-id="e72f9-108">Office hosts require that add-ins initialize within 5 seconds of activation.</span></span> <span data-ttu-id="e72f9-109">このしきい値内にアドインがアクティブにならない場合は、応答なしが宣言され、エラー メッセージがユーザーに表示されます。</span><span class="sxs-lookup"><span data-stu-id="e72f9-109">If your add-in doesn't activate within this threshold, it will be declared unresponsive and an error message will be displayed to the user.</span></span>

## <a name="api-versioning-and-backward-compatibility"></a><span data-ttu-id="e72f9-110">API のバージョン管理と下位互換性</span><span class="sxs-lookup"><span data-stu-id="e72f9-110">API versioning and backward compatibility</span></span>

<span data-ttu-id="e72f9-111">前の HTML スニペットで、 `/1/` CDN URL の先頭にある、 `office.js` バージョン1の Office .js で最新の増分リリースを指定します。</span><span class="sxs-lookup"><span data-stu-id="e72f9-111">In the previous HTML snippet, the `/1/` in front of `office.js` in the CDN URL specifies the latest incremental release within version 1 of Office.js.</span></span> <span data-ttu-id="e72f9-112">Office JavaScript API は下位互換性を維持しているため、最新のリリースでは、以前のバージョン1で導入された API メンバーを引き続きサポートしています。</span><span class="sxs-lookup"><span data-stu-id="e72f9-112">Because the Office JavaScript API maintains backward compatibility, the latest release will continue to support API members that were introduced earlier in version 1.</span></span> <span data-ttu-id="e72f9-113">既存のプロジェクトをアップグレードする必要がある場合は、「 [Office JAVASCRIPT API およびマニフェストスキーマファイルのバージョンを更新](update-your-javascript-api-for-office-and-manifest-schema-version.md)する」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="e72f9-113">If you need to upgrade an existing project, see [Update the version of your Office JavaScript API and manifest schema files](update-your-javascript-api-for-office-and-manifest-schema-version.md).</span></span> 

<span data-ttu-id="e72f9-p104">AppSource から Office アドインを発行する場合は、この CDN の参照を使用する必要があります。ローカル参照は、内部シナリオ、開発シナリオ、デバッグ シナリオにのみ適用できます。</span><span class="sxs-lookup"><span data-stu-id="e72f9-p104">If you plan to publish your Office Add-in from AppSource, you must use this CDN reference. Local references are only appropriate for internal, development, and debugging scenarios.</span></span>

> [!NOTE]
> <span data-ttu-id="e72f9-116">プレビュー API を使用するには、CDN (`https://appsforoffice.microsoft.com/lib/beta/hosted/office.js`) にある Office JavaScript API ライブラリのプレビュー バージョンを参照します。</span><span class="sxs-lookup"><span data-stu-id="e72f9-116">To use preview APIs, reference the preview version of the Office JavaScript API library on the CDN: `https://appsforoffice.microsoft.com/lib/beta/hosted/office.js`.</span></span>

## <a name="enabling-intellisense-for-a-typescript-project"></a><span data-ttu-id="e72f9-117">TypeScript プロジェクトに対して Intellisense を有効にする</span><span class="sxs-lookup"><span data-stu-id="e72f9-117">Enabling Intellisense for a TypeScript project</span></span>

<span data-ttu-id="e72f9-118">前述したように Office JavaScript API を参照するだけでなく、[指定](https://github.com/DefinitelyTyped/DefinitelyTyped/tree/master/types/office-js)した型定義を使用して TypeScript アドインプロジェクトの Intellisense を有効にすることもできます。</span><span class="sxs-lookup"><span data-stu-id="e72f9-118">In addition to referencing the Office JavaScript API as described previously, you can also enable Intellisense for TypeScript add-in project by using the type definitions from [DefinitelyTyped](https://github.com/DefinitelyTyped/DefinitelyTyped/tree/master/types/office-js).</span></span> <span data-ttu-id="e72f9-119">これを行うには、プロジェクトフォルダーのルートから、ノードが有効なシステムプロンプト (または git bash ウィンドウ) で次のコマンドを実行します。</span><span class="sxs-lookup"><span data-stu-id="e72f9-119">To do so, run the following command in a Node-enabled system prompt (or git bash window) from the root of your project folder.</span></span> <span data-ttu-id="e72f9-120">(npm を含む) [Node.js](https://nodejs.org) をインストールしておく必要があります。</span><span class="sxs-lookup"><span data-stu-id="e72f9-120">You must have [Node.js](https://nodejs.org) installed (which includes npm).</span></span>

```command&nbsp;line
npm install --save-dev @types/office-js
```

> [!NOTE]
> <span data-ttu-id="e72f9-121">プレビュー Api に対して Intellisense を有効にするには、プロジェクトフォルダーのルートで次のコマンドを実行することによって[、型定義](https://github.com/DefinitelyTyped/DefinitelyTyped/tree/master/types/office-js-preview)のプレビュータイプ定義を使用します。</span><span class="sxs-lookup"><span data-stu-id="e72f9-121">To enable Intellisense for preview APIs, use the preview type definitions from [DefinitelyTyped](https://github.com/DefinitelyTyped/DefinitelyTyped/tree/master/types/office-js-preview) by running the following command in the root of your project folder:</span></span> 
>
> `npm install --save-dev @types/office-js-preview`

## <a name="see-also"></a><span data-ttu-id="e72f9-122">関連項目</span><span class="sxs-lookup"><span data-stu-id="e72f9-122">See also</span></span>

- [<span data-ttu-id="e72f9-123">Office JavaScript API について</span><span class="sxs-lookup"><span data-stu-id="e72f9-123">Understanding the Office JavaScript API</span></span>](understanding-the-javascript-api-for-office.md)
- [<span data-ttu-id="e72f9-124">Office の JavaScript API</span><span class="sxs-lookup"><span data-stu-id="e72f9-124">Office JavaScript API</span></span>](../reference/javascript-api-for-office.md)
