---
title: Office JavaScript API ライブラリの参照
description: アドインで JavaScript API ライブラリOfficeタイプ定義を参照する方法について説明します。
ms.date: 02/18/2021
localization_priority: Normal
ms.openlocfilehash: 04f97412c07cb39f5b2f753c3ce14e56e87c3de5
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349757"
---
# <a name="referencing-the-office-javascript-api-library"></a><span data-ttu-id="55c10-103">Office JavaScript API ライブラリの参照</span><span class="sxs-lookup"><span data-stu-id="55c10-103">Referencing the Office JavaScript API library</span></span>

<span data-ttu-id="55c10-104">[JavaScript API Officeには](../reference/javascript-api-for-office.md)、アドインがアプリケーションと対話するために使用できる API がOfficeされます。</span><span class="sxs-lookup"><span data-stu-id="55c10-104">The [Office JavaScript API](../reference/javascript-api-for-office.md) library provides the APIs that your add-in can use to interact with the Office application.</span></span> <span data-ttu-id="55c10-105">ライブラリを参照する最も簡単な方法は、HTML ページのセクション内にCDNタグを追加してコンテンツ配信ネットワーク (CDN) `<script>` `<head>` を使用する方法です。</span><span class="sxs-lookup"><span data-stu-id="55c10-105">The simplest way to reference the library is to use the content delivery network (CDN) by adding the following `<script>` tag within the `<head>` section of your HTML page.</span></span>

```html
<head>
    ...
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
</head>
```

<span data-ttu-id="55c10-106">これにより、Office JavaScript API ファイルが初めて読み込まれると、Office.js の最新の実装と指定したバージョンの関連ファイルが使用されます。</span><span class="sxs-lookup"><span data-stu-id="55c10-106">This will download and cache the Office JavaScript API files the first time your add-in loads to make sure that it is using the most up-to-date implementation of Office.js and its associated files for the specified version.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="55c10-107">ページのセクション内Office JavaScript API を参照して、本文要素の前に API が完全に初期化 `<head>` される必要があります。</span><span class="sxs-lookup"><span data-stu-id="55c10-107">You must reference the Office JavaScript API from inside the `<head>` section of the page to ensure that the API is fully initialized prior to any body elements.</span></span>

## <a name="api-versioning-and-backward-compatibility"></a><span data-ttu-id="55c10-108">API のバージョン管理と下位互換性</span><span class="sxs-lookup"><span data-stu-id="55c10-108">API versioning and backward compatibility</span></span>

<span data-ttu-id="55c10-109">前の HTML スニペットでは、URL の前に、CDN バージョン 1 の増分リリース `/1/` `office.js` を指定Office.js。</span><span class="sxs-lookup"><span data-stu-id="55c10-109">In the previous HTML snippet, the `/1/` in front of `office.js` in the CDN URL specifies the latest incremental release within version 1 of Office.js.</span></span> <span data-ttu-id="55c10-110">JavaScript API Office互換性が維持されるので、最新のリリースでは、バージョン 1 で以前に導入された API メンバーを引き続きサポートします。</span><span class="sxs-lookup"><span data-stu-id="55c10-110">Because the Office JavaScript API maintains backward compatibility, the latest release will continue to support API members that were introduced earlier in version 1.</span></span> <span data-ttu-id="55c10-111">既存のプロジェクトをアップグレードする必要がある場合は[、「JavaScript API](update-your-javascript-api-for-office-and-manifest-schema-version.md)とマニフェスト スキーマ ファイルのバージョンOffice更新する」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="55c10-111">If you need to upgrade an existing project, see [Update the version of your Office JavaScript API and manifest schema files](update-your-javascript-api-for-office-and-manifest-schema-version.md).</span></span> 

<span data-ttu-id="55c10-p103">AppSource から Office アドインを発行する場合は、この CDN の参照を使用する必要があります。ローカル参照は、内部シナリオ、開発シナリオ、デバッグ シナリオにのみ適用できます。</span><span class="sxs-lookup"><span data-stu-id="55c10-p103">If you plan to publish your Office Add-in from AppSource, you must use this CDN reference. Local references are only appropriate for internal, development, and debugging scenarios.</span></span>

> [!NOTE]
> <span data-ttu-id="55c10-114">プレビュー API を使用するには、CDN (`https://appsforoffice.microsoft.com/lib/beta/hosted/office.js`) にある Office JavaScript API ライブラリのプレビュー バージョンを参照します。</span><span class="sxs-lookup"><span data-stu-id="55c10-114">To use preview APIs, reference the preview version of the Office JavaScript API library on the CDN: `https://appsforoffice.microsoft.com/lib/beta/hosted/office.js`.</span></span>

## <a name="enabling-intellisense-for-a-typescript-project"></a><span data-ttu-id="55c10-115">TypeScript プロジェクトIntelliSenseを有効にする</span><span class="sxs-lookup"><span data-stu-id="55c10-115">Enabling IntelliSense for a TypeScript project</span></span>

<span data-ttu-id="55c10-116">前述のように Office JavaScript API を参照する以外に[、DefinitelyTyped](https://github.com/DefinitelyTyped/DefinitelyTyped/tree/master/types/office-js)の型定義を使用して TypeScript アドイン プロジェクトの IntelliSense を有効にすることもできます。</span><span class="sxs-lookup"><span data-stu-id="55c10-116">In addition to referencing the Office JavaScript API as described previously, you can also enable IntelliSense for TypeScript add-in project by using the type definitions from [DefinitelyTyped](https://github.com/DefinitelyTyped/DefinitelyTyped/tree/master/types/office-js).</span></span> <span data-ttu-id="55c10-117">これを行うには、プロジェクト フォルダーのルートからノード対応のシステム プロンプト (または git bash ウィンドウ) で次のコマンドを実行します。</span><span class="sxs-lookup"><span data-stu-id="55c10-117">To do so, run the following command in a Node-enabled system prompt (or git bash window) from the root of your project folder.</span></span> <span data-ttu-id="55c10-118">(npm を含む) [Node.js](https://nodejs.org) をインストールしておく必要があります。</span><span class="sxs-lookup"><span data-stu-id="55c10-118">You must have [Node.js](https://nodejs.org) installed (which includes npm).</span></span>

```command&nbsp;line
npm install --save-dev @types/office-js
```

## <a name="preview-apis"></a><span data-ttu-id="55c10-119">プレビュー API</span><span class="sxs-lookup"><span data-stu-id="55c10-119">Preview APIs</span></span>

<span data-ttu-id="55c10-120">新しい JavaScript API は、最初に "プレビュー" で導入され、後で十分なテストが行われるとユーザーフィードバックが必要になった後、特定の番号付き要件セットの一部になります。</span><span class="sxs-lookup"><span data-stu-id="55c10-120">New JavaScript APIs are first introduced in "preview" and later become part of a specific numbered requirement set after sufficient testing occurs and user feedback is required.</span></span>

[!INCLUDE [Information about using preview APIs](../includes/using-preview-apis-host.md)]

## <a name="see-also"></a><span data-ttu-id="55c10-121">関連項目</span><span class="sxs-lookup"><span data-stu-id="55c10-121">See also</span></span>

- [<span data-ttu-id="55c10-122">Office JavaScript API について</span><span class="sxs-lookup"><span data-stu-id="55c10-122">Understanding the Office JavaScript API</span></span>](understanding-the-javascript-api-for-office.md)
- [<span data-ttu-id="55c10-123">Office の JavaScript API</span><span class="sxs-lookup"><span data-stu-id="55c10-123">Office JavaScript API</span></span>](../reference/javascript-api-for-office.md)
