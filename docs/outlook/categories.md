---
title: カテゴリの取得と設定
description: '[方法] メールボックスとアイテムのカテゴリを管理する'
ms.date: 01/14/2020
localization_priority: Normal
ms.openlocfilehash: d4589571de47218741308c01caec0166d72919d8
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608979"
---
# <a name="get-and-set-categories"></a><span data-ttu-id="14d73-103">カテゴリの取得と設定</span><span class="sxs-lookup"><span data-stu-id="14d73-103">Get and set categories</span></span>

<span data-ttu-id="14d73-104">Outlook では、ユーザーはメールボックスのデータを整理する手段として、メッセージや予定に分類項目を適用することができます。</span><span class="sxs-lookup"><span data-stu-id="14d73-104">In Outlook, a user can apply categories to messages and appointments as a means of organizing their mailbox data.</span></span> <span data-ttu-id="14d73-105">ユーザーは、自分のメールボックスの色分けされたカテゴリのマスターリストを定義し、そのうちの1つまたは複数のカテゴリを任意のメッセージアイテムまたは予定アイテムに適用することができます。</span><span class="sxs-lookup"><span data-stu-id="14d73-105">The user defines the master list of color-coded categories for their mailbox, and can then apply one or more of those categories to any message or appointment item.</span></span> <span data-ttu-id="14d73-106">マスターリストの各[カテゴリ](/javascript/api/outlook/office.categorydetails)は、ユーザーが指定した名前と[色](/javascript/api/outlook/office.mailboxenums.categorycolor)で表されます。</span><span class="sxs-lookup"><span data-stu-id="14d73-106">Each [category](/javascript/api/outlook/office.categorydetails) in the master list is represented by the name and [color](/javascript/api/outlook/office.mailboxenums.categorycolor) that the user specifies.</span></span> <span data-ttu-id="14d73-107">Office JavaScript API を使用して、メールボックスのカテゴリマスターリストとアイテムに適用されるカテゴリを管理できます。</span><span class="sxs-lookup"><span data-stu-id="14d73-107">You can use the Office JavaScript API to manage the categories master list on the mailbox and the categories applied to an item.</span></span>

> [!NOTE]
> <span data-ttu-id="14d73-108">この機能のサポートは、要件セット1.8 で導入されました。</span><span class="sxs-lookup"><span data-stu-id="14d73-108">Support for this feature was introduced in requirement set 1.8.</span></span> <span data-ttu-id="14d73-109">この要件セットをサポートする [クライアントおよびプラットフォーム](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) を参照してください。</span><span class="sxs-lookup"><span data-stu-id="14d73-109">See [clients and platforms](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) that support this requirement set.</span></span>

## <a name="manage-categories-in-the-master-list"></a><span data-ttu-id="14d73-110">マスターリストでカテゴリを管理する</span><span class="sxs-lookup"><span data-stu-id="14d73-110">Manage categories in the master list</span></span>

<span data-ttu-id="14d73-111">メールボックスのマスターリストにあるカテゴリのみが、メッセージまたは予定に適用できます。</span><span class="sxs-lookup"><span data-stu-id="14d73-111">Only categories in the master list on your mailbox are available for you to apply to a message or appointment.</span></span> <span data-ttu-id="14d73-112">この API を使用して、マスターカテゴリの追加、取得、および削除を行うことができます。</span><span class="sxs-lookup"><span data-stu-id="14d73-112">You can use the API to add, get, and remove master categories.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="14d73-113">このアドインでカテゴリマスターリストを管理するには、マニフェスト内のノードをに設定する必要があり `Permissions` `ReadWriteMailbox` ます。</span><span class="sxs-lookup"><span data-stu-id="14d73-113">For the add-in to manage the categories master list, you must set the `Permissions` node in the manifest to `ReadWriteMailbox`.</span></span>

### <a name="add-master-categories"></a><span data-ttu-id="14d73-114">マスターカテゴリを追加する</span><span class="sxs-lookup"><span data-stu-id="14d73-114">Add master categories</span></span>

<span data-ttu-id="14d73-115">次の例は、"至急!" という名前の分類項目を追加する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="14d73-115">The following example shows how to add a category named "Urgent!"</span></span> <span data-ttu-id="14d73-116">をマスターリストに追加するには、[メールボックス. masterCategories](/javascript/api/outlook/office.mailbox#mastercategories)で[addasync](/javascript/api/outlook/office.mastercategories#addasync-categories--options--callback-)を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="14d73-116">to the master list by calling [addAsync](/javascript/api/outlook/office.mastercategories#addasync-categories--options--callback-) on [mailbox.masterCategories](/javascript/api/outlook/office.mailbox#mastercategories).</span></span>

```js
var masterCategoriesToAdd = [
    {
        "displayName": "Urgent!",
        "color": Office.MailboxEnums.CategoryColor.Preset0
    }
];

Office.context.mailbox.masterCategories.addAsync(masterCategoriesToAdd, function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        console.log("Successfully added categories to master list");
    } else {
        console.log("masterCategories.addAsync call failed with error: " + asyncResult.error.message);
    }
});
```

### <a name="get-master-categories"></a><span data-ttu-id="14d73-117">マスターカテゴリを取得する</span><span class="sxs-lookup"><span data-stu-id="14d73-117">Get master categories</span></span>

<span data-ttu-id="14d73-118">次の例は、 [getAsync](/javascript/api/outlook/office.mastercategories#getasync-options--callback-)の[メールボックス. mastercategories](/javascript/api/outlook/office.mailbox#mastercategories)で、カテゴリの一覧を取得する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="14d73-118">The following example shows how to get the list of categories by calling [getAsync](/javascript/api/outlook/office.mastercategories#getasync-options--callback-) on [mailbox.masterCategories](/javascript/api/outlook/office.mailbox#mastercategories).</span></span>

```js
Office.context.mailbox.masterCategories.getAsync(function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.log("Action failed with error: " + asyncResult.error.message);
    } else {
        var masterCategories = asyncResult.value;
        console.log("Master categories:");
        masterCategories.forEach(function (item) {
            console.log("-- " + JSON.stringify(item));
        });
    }
});
```

### <a name="remove-master-categories"></a><span data-ttu-id="14d73-119">マスターシェイプカテゴリを削除する</span><span class="sxs-lookup"><span data-stu-id="14d73-119">Remove master categories</span></span>

<span data-ttu-id="14d73-120">次の例は、"至急!" という名前の分類項目を削除する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="14d73-120">The following example shows how to remove the category named "Urgent!"</span></span> <span data-ttu-id="14d73-121">マスターリストから、RemoveAsync[カテゴリ](/javascript/api/outlook/office.mailbox#mastercategories)で [ [removeAsync](/javascript/api/outlook/office.mastercategories#removeasync-categories--options--callback-) ] を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="14d73-121">from the master list by calling [removeAsync](/javascript/api/outlook/office.mastercategories#removeasync-categories--options--callback-) on [mailbox.masterCategories](/javascript/api/outlook/office.mailbox#mastercategories).</span></span>

```js
var masterCategoriesToRemove = ["Urgent!"];

Office.context.mailbox.masterCategories.removeAsync(masterCategoriesToRemove, function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        console.log("Successfully removed categories from master list");
    } else {
        console.log("masterCategories.removeAsync call failed with error: " + asyncResult.error.message);
    }
});
```

## <a name="manage-categories-on-a-message-or-appointment"></a><span data-ttu-id="14d73-122">メッセージまたは予定の分類項目を管理する</span><span class="sxs-lookup"><span data-stu-id="14d73-122">Manage categories on a message or appointment</span></span>

<span data-ttu-id="14d73-123">API を使用して、メッセージアイテムまたは予定アイテムの分類項目の追加、取得、削除を行うことができます。</span><span class="sxs-lookup"><span data-stu-id="14d73-123">You can use the API to add, get, and remove categories for a message or appointment item.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="14d73-124">メールボックスのマスターリストにあるカテゴリのみが、メッセージまたは予定に適用できます。</span><span class="sxs-lookup"><span data-stu-id="14d73-124">Only categories in the master list on your mailbox are available for you to apply to a message or appointment.</span></span> <span data-ttu-id="14d73-125">詳細については、前のセクション「[マスターリストでカテゴリを管理](#manage-categories-in-the-master-list)する」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="14d73-125">See the earlier section [Manage categories in the master list](#manage-categories-in-the-master-list) for more information.</span></span>
>
> <span data-ttu-id="14d73-126">Web 上の Outlook では、API を使用してメッセージのカテゴリを閲覧モードで管理することはできません。</span><span class="sxs-lookup"><span data-stu-id="14d73-126">In Outlook on the web, you can't use the API to manage categories on a message in Read mode.</span></span>

### <a name="add-categories-to-an-item"></a><span data-ttu-id="14d73-127">アイテムに分類項目を追加する</span><span class="sxs-lookup"><span data-stu-id="14d73-127">Add categories to an item</span></span>

<span data-ttu-id="14d73-128">次の例は、"至急!" という名前の分類項目を適用する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="14d73-128">The following example shows how to apply the category named "Urgent!"</span></span> <span data-ttu-id="14d73-129">で[Addasync](/javascript/api/outlook/office.categories#addasync-categories--options--callback-)を呼び出して、現在のアイテムに追加 `item.categories` します。</span><span class="sxs-lookup"><span data-stu-id="14d73-129">to the current item by calling [addAsync](/javascript/api/outlook/office.categories#addasync-categories--options--callback-) on `item.categories`.</span></span>

```js
var categoriesToAdd = ["Urgent!"];

Office.context.mailbox.item.categories.addAsync(categoriesToAdd, function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        console.log("Successfully added categories");
    } else {
        console.log("categories.addAsync call failed with error: " + asyncResult.error.message);
    }
});
```

### <a name="get-an-items-categories"></a><span data-ttu-id="14d73-130">アイテムのカテゴリを取得する</span><span class="sxs-lookup"><span data-stu-id="14d73-130">Get an item's categories</span></span>

<span data-ttu-id="14d73-131">次の例は、 [getAsync](/javascript/api/outlook/office.categories#getasync-options--callback-) on を呼び出すことによって、現在のアイテムに適用されているカテゴリを取得する方法を示して `item.categories` います。</span><span class="sxs-lookup"><span data-stu-id="14d73-131">The following example shows how to get the categories applied to the current item by calling [getAsync](/javascript/api/outlook/office.categories#getasync-options--callback-) on `item.categories`.</span></span>

```js
Office.context.mailbox.item.categories.getAsync(function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.log("Action failed with error: " + asyncResult.error.message);
    } else {
        var categories = asyncResult.value;
        console.log("Categories:");
        categories.forEach(function (item) {
            console.log("-- " + JSON.stringify(item));
        });
    }
});
```

### <a name="remove-categories-from-an-item"></a><span data-ttu-id="14d73-132">アイテムからカテゴリを削除する</span><span class="sxs-lookup"><span data-stu-id="14d73-132">Remove categories from an item</span></span>

<span data-ttu-id="14d73-133">次の例は、"至急!" という名前の分類項目を削除する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="14d73-133">The following example shows how to remove the category named "Urgent!"</span></span> <span data-ttu-id="14d73-134">現在のアイテムから[removeAsync](/javascript/api/outlook/office.categories#removeasync-categories--options--callback-) on を呼び出し `item.categories` ます。</span><span class="sxs-lookup"><span data-stu-id="14d73-134">from the current item by calling [removeAsync](/javascript/api/outlook/office.categories#removeasync-categories--options--callback-) on `item.categories`.</span></span>

```js
var categoriesToRemove = ["Urgent!"];

Office.context.mailbox.item.categories.removeAsync(categoriesToRemove, function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        console.log("Successfully removed categories");
    } else {
        console.log("categories.removeAsync call failed with error: " + asyncResult.error.message);
    }
});
```

## <a name="see-also"></a><span data-ttu-id="14d73-135">関連項目</span><span class="sxs-lookup"><span data-stu-id="14d73-135">See also</span></span>

- [<span data-ttu-id="14d73-136">Outlook のアクセス許可</span><span class="sxs-lookup"><span data-stu-id="14d73-136">Outlook permissions</span></span>](understanding-outlook-add-in-permissions.md)
- [<span data-ttu-id="14d73-137">マニフェストの Permissions 要素</span><span class="sxs-lookup"><span data-stu-id="14d73-137">Permissions element in the manifest</span></span>](../reference/manifest/permissions.md)
