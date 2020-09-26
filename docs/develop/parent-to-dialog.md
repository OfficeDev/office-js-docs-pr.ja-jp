---
title: ホストページからダイアログボックスにメッセージを渡す代替方法
description: MessageChild メソッドがサポートされていない場合に使用する回避策について説明します。
ms.date: 09/24/2020
localization_priority: Normal
ms.openlocfilehash: 8f44f7f5c145b58d13e7387d01e28fd349a512fc
ms.sourcegitcommit: b47318a24a50443b0579e05e178b3bb5433c372f
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/25/2020
ms.locfileid: "48279483"
---
# <a name="alternative-ways-of-passing-messages-to-a-dialog-box-from-its-host-page"></a><span data-ttu-id="12496-103">ホストページからダイアログボックスにメッセージを渡す代替方法</span><span class="sxs-lookup"><span data-stu-id="12496-103">Alternative ways of passing messages to a dialog box from its host page</span></span>

<span data-ttu-id="12496-104">親ページから子ダイアログボックスにデータおよびメッセージを渡す方法としては、 `messageChild` 「 [office アドインで OFFICE ダイアログ API を使用](dialog-api-in-office-add-ins.md#pass-information-to-the-dialog-box)する」で説明されている方法を使用することをお勧めします。Add-in [api 1.2 要件セット](../reference/requirement-sets/dialog-api-requirement-sets.md)をサポートしていないプラットフォームまたはホストでアドインが実行されている場合は、次の2つの方法で情報をダイアログボックスに渡すことができます。</span><span class="sxs-lookup"><span data-stu-id="12496-104">The recommended way to pass data and messages from a parent page to a child dialog box is with the `messageChild` method as described in [Use the Office dialog API in your Office Add-ins](dialog-api-in-office-add-ins.md#pass-information-to-the-dialog-box). If your add-in is running on a platform or host that does not support the [DialogApi 1.2 requirement set](../reference/requirement-sets/dialog-api-requirement-sets.md), there are two other ways that you can pass information to the dialog box:</span></span>

- <span data-ttu-id="12496-105">`displayDialogAsync` に渡される URL にクエリ パラメーターを追加します。</span><span class="sxs-lookup"><span data-stu-id="12496-105">Add query parameters to the URL that is passed to `displayDialogAsync`.</span></span>
- <span data-ttu-id="12496-106">ホスト ウィンドウとダイアログ ボックスの両方にアクセス可能な場所に情報を格納します。</span><span class="sxs-lookup"><span data-stu-id="12496-106">Store the information somewhere that is accessible to both the host window and dialog box.</span></span> <span data-ttu-id="12496-107">2つのウィンドウは、共通のセッション記憶域 ( [Window. sessionstorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage) プロパティ) を共有しませんが、 *同じドメイン* (ポート番号を含む) がある場合は、共通の [ローカル記憶域](https://www.w3schools.com/html/html5_webstorage.asp)を共有します。\*</span><span class="sxs-lookup"><span data-stu-id="12496-107">The two windows do not share a common session storage (the [Window.sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage) property), but *if they have the same domain* (including port number, if any), they share a common [Local Storage](https://www.w3schools.com/html/html5_webstorage.asp).\*</span></span>


> [!NOTE]
> <span data-ttu-id="12496-108">\* トークン処理の戦略に影響を与えるバグがあります。</span><span class="sxs-lookup"><span data-stu-id="12496-108">\* There is a bug that will effect your strategy for token handling.</span></span> <span data-ttu-id="12496-109">Safari または Microsoft Edge ブラウザーの **Office on the web** でアドインを実行している場合、ダイアログ ボックスとタスク ウィンドウは同じローカル ストレージを共有しないため、これらの間の通信に使用できません。</span><span class="sxs-lookup"><span data-stu-id="12496-109">If the add-in is running in **Office on the web** in either the Safari or Edge browser, the dialog box and task pane do not share the same Local Storage, so it cannot be used to communicate between them.</span></span>

## <a name="use-local-storage"></a><span data-ttu-id="12496-110">ローカル ストレージの使用</span><span class="sxs-lookup"><span data-stu-id="12496-110">Use local storage</span></span>

<span data-ttu-id="12496-111">ローカルストレージを使用するには、 `setItem` 次の `window.localStorage` `displayDialogAsync` 例に示すように、呼び出しの前に、ホストページでオブジェクトのメソッドを呼び出します。</span><span class="sxs-lookup"><span data-stu-id="12496-111">To use local storage, call the `setItem` method of the `window.localStorage` object in the host page before the `displayDialogAsync` call, as in the following example:</span></span>

```js
localStorage.setItem("clientID", "15963ac5-314f-4d9b-b5a1-ccb2f1aea248");
```

<span data-ttu-id="12496-112">ダイアログ ボックス内のコードは、次の例に示すように、必要に応じて項目を読み取ります。</span><span class="sxs-lookup"><span data-stu-id="12496-112">Code in the dialog box reads the item when it's needed, as in the following example:</span></span>

```js
var clientID = localStorage.getItem("clientID");
// You can also use property syntax:
// var clientID = localStorage.clientID;
```

## <a name="use-query-parameters"></a><span data-ttu-id="12496-113">クエリ パラメーターの使用</span><span class="sxs-lookup"><span data-stu-id="12496-113">Use query parameters</span></span>

<span data-ttu-id="12496-114">次の例では、クエリ パラメーターを使用してデータを渡す方法を示します。</span><span class="sxs-lookup"><span data-stu-id="12496-114">The following example shows how to pass data with a query parameter:</span></span>

```js
Office.context.ui.displayDialogAsync('https://myAddinDomain/myDialog.html?clientID=15963ac5-314f-4d9b-b5a1-ccb2f1aea248');
```

<span data-ttu-id="12496-115">この手法を使用するサンプルについては、「[PowerPoint アドインで Microsoft Graph を使用した Excel グラフの挿入](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="12496-115">For a sample that uses this technique, see [Insert Excel charts using Microsoft Graph in a PowerPoint add-in](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart).</span></span>

<span data-ttu-id="12496-116">ダイアログ ボックス内のコードは、URL を解析し、パラメーター値を読み取ることができます。</span><span class="sxs-lookup"><span data-stu-id="12496-116">Code in your dialog box can parse the URL and read the parameter value.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="12496-117">Office は、`_host_info` に渡される URL に `displayDialogAsync` というクエリ パラメーターを自動的に追加します (カスタム クエリ パラメーターが存在する場合は、その後に追加されます。</span><span class="sxs-lookup"><span data-stu-id="12496-117">Office automatically adds a query parameter called `_host_info` to the URL that is passed to `displayDialogAsync`.</span></span> <span data-ttu-id="12496-118">ダイアログ ボックスが移動する先の後続の URL には追加されません)。</span><span class="sxs-lookup"><span data-stu-id="12496-118">(It is appended after your custom query parameters, if any.</span></span> <span data-ttu-id="12496-119">Microsoft は、将来、この値の内容を変更したり、完全に削除したりする可能性があるため、コードでこの値の内容を読み取らないでください。</span><span class="sxs-lookup"><span data-stu-id="12496-119">It is not appended to any subsequent URLs that the dialog box navigates to.) Microsoft may change the content of this value, or remove it entirely, in the future, so your code should not read it.</span></span> <span data-ttu-id="12496-120">ダイアログボックスのセッションストレージ ( [sessionstorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage) プロパティ) に同じ値が追加されます。</span><span class="sxs-lookup"><span data-stu-id="12496-120">The same value is added to the dialog box's session storage (the [Window.sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage) property).</span></span> <span data-ttu-id="12496-121">この場合も、*コードではこの値に対する読み取りも書き込みも行わないでください*。</span><span class="sxs-lookup"><span data-stu-id="12496-121">Again, *your code should neither read nor write to this value*.</span></span>
