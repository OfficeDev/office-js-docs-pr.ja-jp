---
title: Outlook アドインの API
description: Outlook アドインの API を参照して、Outlook アドインにアクセス許可を宣言する方法について説明します。
ms.date: 10/31/2019
localization_priority: Normal
ms.openlocfilehash: 2abe365f1606789b1c6ac113b133019055767b28
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/20/2020
ms.locfileid: "42166607"
---
# <a name="outlook-add-in-apis"></a><span data-ttu-id="24033-103">Outlook アドインの API</span><span class="sxs-lookup"><span data-stu-id="24033-103">Outlook add-in APIs</span></span>

<span data-ttu-id="24033-104">Outlook アドインで API を使用するには、Office.js ライブラリの場所、要件セット、スキーマ、アクセス許可を指定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="24033-104">To use APIs in your Outlook add-in, you must specify the location of the Office.js library, the requirement set, the schema, and the permissions.</span></span>

## <a name="officejs-library"></a><span data-ttu-id="24033-105">Office.js ライブラリ</span><span class="sxs-lookup"><span data-stu-id="24033-105">Office.js library</span></span>

<span data-ttu-id="24033-106">Outlook アドイン API と対話操作するには、Office.js の JavaScript API を使用する必要があります。</span><span class="sxs-lookup"><span data-stu-id="24033-106">To interact with the Outlook add-in API, you need to use the JavaScript APIs in Office.js.</span></span> <span data-ttu-id="24033-107">ライブラリ用の CDN は `https://appsforoffice.microsoft.com/lib/1/hosted/Office.js` です。</span><span class="sxs-lookup"><span data-stu-id="24033-107">The CDN for the library is `https://appsforoffice.microsoft.com/lib/1/hosted/Office.js`.</span></span> <span data-ttu-id="24033-108">AppSource に送られるアドインは、この CDN で Office.js を参照しなければなりません。ローカル参照は使用できません。</span><span class="sxs-lookup"><span data-stu-id="24033-108">Add-ins submitted to AppSource must reference Office.js by this CDN; they can't use a local reference.</span></span>

<span data-ttu-id="24033-109">アドインの UI を実行する Web ページ (.html、.aspx、.php のファイル) の `<head>` タグの中の `<script>` タグの中で CDN を参照します。</span><span class="sxs-lookup"><span data-stu-id="24033-109">Reference the CDN in a `<script>` tag in the `<head>` tag of the web page (.html, .aspx, or .php file) that implements the UI of your add-in.</span></span>

```HTML
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
```
<span data-ttu-id="24033-p102">新しい API が追加されても、Office.js への URL は同じままになります。URL 内のバージョンは、既存の API の動作を分割する場合にのみ変更されます。</span><span class="sxs-lookup"><span data-stu-id="24033-p102">As we add new APIs, the URL to Office.js will stay the same. We will change the version in the URL only if we break an existing API behavior.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="24033-112">Office ホスト アプリケーションのアドインを開発する場合は、ページの `<head>` セクションの内側から JavaScript API for Office を参照します。</span><span class="sxs-lookup"><span data-stu-id="24033-112">When developing an add-in for any Office host application, reference the JavaScript API for Office from inside the `<head>` section of the page.</span></span> <span data-ttu-id="24033-113">これにより、あらゆる body 要素の前に API が完全に初期化されます。</span><span class="sxs-lookup"><span data-stu-id="24033-113">This ensures that the API is fully initialized prior to any body elements.</span></span> <span data-ttu-id="24033-114">Office ホストでは、アクティブ化の 5 秒以内にアドインを初期化する必要があります。</span><span class="sxs-lookup"><span data-stu-id="24033-114">Office hosts require that add-ins initialize within 5 seconds of activation.</span></span> <span data-ttu-id="24033-115">このしきい値を超えるとアドインが応答なしと宣言され、ユーザーにエラー メッセージが表示されます。</span><span class="sxs-lookup"><span data-stu-id="24033-115">Crossing this threshold results in the add-in being declared unresponsive and an error message is displayed to the user.</span></span>

## <a name="requirement-sets"></a><span data-ttu-id="24033-116">要件セット</span><span class="sxs-lookup"><span data-stu-id="24033-116">Requirement sets</span></span>

<span data-ttu-id="24033-117">すべての Outlook API は `Mailbox` 要件セットに属しています。</span><span class="sxs-lookup"><span data-stu-id="24033-117">All Outlook APIs belong to the `Mailbox` requirement set.</span></span> <span data-ttu-id="24033-118">`Mailbox` の要件セットにはバージョンがあり、リリースされる新しい API の各セットは、新しいバージョンのセットに属しています。</span><span class="sxs-lookup"><span data-stu-id="24033-118">The `Mailbox` requirement set has versions, and each new set of APIs that are released belongs to a higher version of the set.</span></span> <span data-ttu-id="24033-119">最新の API セットがリリースされても、すべての Outlook クライアントがそれをサポートするわけではありませんが、ある Outlook クライアントが 1 つの要件セットのサポートを宣言した場合、その要件セットの中のすべての API がサポートされます。</span><span class="sxs-lookup"><span data-stu-id="24033-119">Not all Outlook clients will support the newest set of APIs when they are released, but if an Outlook client declares support for a requirement set, it will support all the APIs in that requirement set.</span></span>

<span data-ttu-id="24033-p105">どの Outlook クライアントにアドインを表示するかを制御するには、最小の要件セットのバージョンをマニフェストで指定します。たとえば、要件セットのバージョン 1.3 を指定すると、最小バージョンの 1.3 をサポートしていない Outlook クライアントにはアドインが表示されなくなります。</span><span class="sxs-lookup"><span data-stu-id="24033-p105">To control which Outlook clients the add-in appears in, specify a minimum requirement set version in the manifest. For example, if you specify requirement set version 1.3, the add-in will not show up in any Outlook client that doesn't support a minimum version of 1.3.</span></span>

<span data-ttu-id="24033-p106">要件セットを指定しても、そのバージョンの API にアドインを限定することにはなりません。要件セット v1.1 を指定しているアドインが、v1.3 をサポートする Outlook クライアントで実行されると、そのアドインは v1.3 の API を使用できます。要件セットでは、どの Outlook クライアントにアドインを表示するかのみを制御します。</span><span class="sxs-lookup"><span data-stu-id="24033-p106">Specifying a requirement set doesn't limit your add-in to the APIs in that version. If the add-in specifies requirement set v1.1 but is running in an Outlook client that supports v1.3, the add-in can still use v1.3 APIs. The requirement set only controls which Outlook clients the add-in appears in.</span></span>

<span data-ttu-id="24033-125">マニフェストで指定した要件セットよりも上位の要件セットの API が使用できるかどうかを確認する場合は、標準の JavaScript を使用できます。</span><span class="sxs-lookup"><span data-stu-id="24033-125">To check the availability of any APIs from a requirement set greater than the one specified in the manifest, you can use standard JavaScript:</span></span>

```js
if (item.somePropertyOrFunction) {
   item.somePropertyOrFunction...  
}
```

> [!NOTE]
> <span data-ttu-id="24033-126">このような確認は、マニフェストで指定された要件セットのバージョンに存在する API には必要ありません。</span><span class="sxs-lookup"><span data-stu-id="24033-126">These checks are not needed for any APIs that are in the requirement set version specified in the manifest.</span></span>

<span data-ttu-id="24033-127">それなしではアドインの機能が機能しないような、シナリオに絶対必要な API のセットをサポートする最低限要件セットを指定します。</span><span class="sxs-lookup"><span data-stu-id="24033-127">Specify the minimum requirement set that supports the critical set of APIs for your scenario, without which features of your add-in won't work.</span></span> <span data-ttu-id="24033-128">要件セットは、`<Requirements>` 要素内のマニフェストで指定します。</span><span class="sxs-lookup"><span data-stu-id="24033-128">You specify the requirement set in the manifest in the `<Requirements>` element.</span></span> <span data-ttu-id="24033-129">詳細は、[Outlook のアドイン マニフェスト](manifests.md)と「[Outlook API 要件セットについて](../reference/requirement-sets/outlook-api-requirement-sets.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="24033-129">For more information, see [Outlook add-in manifests](manifests.md) and [Understanding Outlook API requirement sets](../reference/requirement-sets/outlook-api-requirement-sets.md).</span></span>

<span data-ttu-id="24033-130">`<Methods>` 要素は Outlook アドインには適用されないので、特定のメソッドのサポートは宣言できません。</span><span class="sxs-lookup"><span data-stu-id="24033-130">The `<Methods>` element doesn't apply to Outlook add-ins, so you can't declare support for specific methods.</span></span>

## <a name="permissions"></a><span data-ttu-id="24033-131">アクセス許可</span><span class="sxs-lookup"><span data-stu-id="24033-131">Permissions</span></span>

<span data-ttu-id="24033-p108">アドインには、そのアドインが必要とする API を使用するための適切なアクセス許可が必要になります。アクセス許可には、4 つのレベルがあります。詳細については、「[Outlook アドインのアクセス許可モデルを理解する](understanding-outlook-add-in-permissions.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="24033-p108">Your add-in requires the appropriate permissions to use the APIs that it needs. There are four levels of permissions. For more details, see [Understanding Outlook add-in permissions](understanding-outlook-add-in-permissions.md).</span></span>

<br/>

|<span data-ttu-id="24033-135">権限レベル</span><span class="sxs-lookup"><span data-stu-id="24033-135">Permission level</span></span>|<span data-ttu-id="24033-136">説明</span><span class="sxs-lookup"><span data-stu-id="24033-136">Description</span></span>|
|:-----|:-----|
| <span data-ttu-id="24033-137">**制限付き**</span><span class="sxs-lookup"><span data-stu-id="24033-137">**Restricted**</span></span> | <span data-ttu-id="24033-138">エンティティは使用できますが、正規表現は使用できません。</span><span class="sxs-lookup"><span data-stu-id="24033-138">Allows use of entities but not regular expressions.</span></span> |
| <span data-ttu-id="24033-139">**アイテムの読み取り**</span><span class="sxs-lookup"><span data-stu-id="24033-139">**Read item**</span></span> | <span data-ttu-id="24033-140">**制限付き**で許可されているものに加えて、以下のものが許可されます。</span><span class="sxs-lookup"><span data-stu-id="24033-140">In addition to what is allowed in **Restricted**, it allows:</span></span><ul><li><span data-ttu-id="24033-141">正規表現</span><span class="sxs-lookup"><span data-stu-id="24033-141">regular expressions</span></span></li><li><span data-ttu-id="24033-142">Outlook アドイン API の読み取りアクセス</span><span class="sxs-lookup"><span data-stu-id="24033-142">Outlook add-in API read access</span></span></li><li><span data-ttu-id="24033-143">アイテムのプロパティとコールバック トークンの取得</span><span class="sxs-lookup"><span data-stu-id="24033-143">getting the item properties and the callback token</span></span></li></ul> |
| <span data-ttu-id="24033-144">**読み取り/書き込み**</span><span class="sxs-lookup"><span data-stu-id="24033-144">**Read/write**</span></span> | <span data-ttu-id="24033-145">**アイテムの読み取り**で許可される内容に加えて、次に示す内容が許可されます。</span><span class="sxs-lookup"><span data-stu-id="24033-145">In addition to what is allowed in **Read item**, it allows:</span></span><ul><li><span data-ttu-id="24033-146">`makeEwsRequestAsync` を除いた、完全な Outlook アドイン API のアクセス</span><span class="sxs-lookup"><span data-stu-id="24033-146">full Outlook add-in API access except `makeEwsRequestAsync`</span></span></li><li><span data-ttu-id="24033-147">アイテムのプロパティの設定</span><span class="sxs-lookup"><span data-stu-id="24033-147">setting the item properties</span></span></li></ul> |
| <span data-ttu-id="24033-148">**メールボックスの読み取り/書き込み**</span><span class="sxs-lookup"><span data-stu-id="24033-148">**Read/write mailbox**</span></span> | <span data-ttu-id="24033-149">**読み取り/書き込み**で許可されているものに加えて、以下のものが許可されます。</span><span class="sxs-lookup"><span data-stu-id="24033-149">In addition to what is allowed in **Read/write**, it allows:</span></span><ul><li><span data-ttu-id="24033-150">アイテムやフォルダーの作成、読み取り、書き込み</span><span class="sxs-lookup"><span data-stu-id="24033-150">creating, reading, writing items and folders</span></span></li><li><span data-ttu-id="24033-151">アイテムの送信</span><span class="sxs-lookup"><span data-stu-id="24033-151">sending items</span></span></li><li><span data-ttu-id="24033-152">[makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) の呼び出し</span><span class="sxs-lookup"><span data-stu-id="24033-152">calling [makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)</span></span></li></ul> |

<span data-ttu-id="24033-153">一般的には、アドインに必要な最低限のアクセス許可を指定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="24033-153">In general, you should specify the minimum permission needed for your add-in.</span></span> <span data-ttu-id="24033-154">アクセス許可は、マニフェスト内の `<Permissions>` 要素で宣言されます。</span><span class="sxs-lookup"><span data-stu-id="24033-154">Permissions are declared in the `<Permissions>` element in the manifest.</span></span> <span data-ttu-id="24033-155">詳細については、「[Outlook アドインのマニフェスト](manifests.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="24033-155">For more information, see [Outlook add-in manifests](manifests.md).</span></span> <span data-ttu-id="24033-156">セキュリティの問題の詳細については、「 [Office アドインのプライバシーとセキュリティ](../concepts/privacy-and-security.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="24033-156">For information about security issues, see [Privacy and security for Office Add-ins](../concepts/privacy-and-security.md).</span></span>


## <a name="see-also"></a><span data-ttu-id="24033-157">関連項目</span><span class="sxs-lookup"><span data-stu-id="24033-157">See also</span></span>

- [<span data-ttu-id="24033-158">Outlook アドインのマニフェスト</span><span class="sxs-lookup"><span data-stu-id="24033-158">Outlook add-in manifests</span></span>](manifests.md)
- [<span data-ttu-id="24033-159">Outlook API 要件セットについて</span><span class="sxs-lookup"><span data-stu-id="24033-159">Understanding Outlook API requirement sets</span></span>](../reference/requirement-sets/outlook-api-requirement-sets.md)
- [<span data-ttu-id="24033-160">Office アドインのプライバシーとセキュリティ</span><span class="sxs-lookup"><span data-stu-id="24033-160">Privacy and security for Office Add-ins</span></span>](../concepts/privacy-and-security.md)
