---
title: カテゴリの取得と設定
description: メールボックスとアイテムのカテゴリを管理する方法。
ms.date: 01/14/2020
ms.localizationpriority: medium
ms.openlocfilehash: 82e6403ad0ac46cd713b9617c089cd4a3884789a
ms.sourcegitcommit: 287a58de82a09deeef794c2aa4f32280efbbe54a
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/28/2022
ms.locfileid: "64496888"
---
# <a name="get-and-set-categories"></a>カテゴリの取得と設定

このOutlookユーザーは、メールボックス データを整理する手段として、メッセージや予定にカテゴリを適用できます。 ユーザーは、自分のメールボックスの色分けされたカテゴリのマスター リストを定義し、それらのカテゴリの 1 つ以上を任意のメッセージまたは予定アイテムに適用できます。 マスター [リスト](/javascript/api/outlook/office.categorydetails)内の各カテゴリは、ユーザーが指定した名前[](/javascript/api/outlook/office.mailboxenums.categorycolor)と色で表されます。 JavaScript API の Officeを使用して、メールボックスのカテゴリ マスター リストとアイテムに適用されるカテゴリを管理できます。

> [!NOTE]
> この機能のサポートは、要件セット 1.8 で導入されました。 この要件セットをサポートする [クライアントおよびプラットフォーム](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets#requirement-sets-supported-by-exchange-servers-and-outlook-clients) を参照してください。

## <a name="manage-categories-in-the-master-list"></a>マスター リストのカテゴリを管理する

メッセージまたは予定に適用できるのは、メールボックスのマスター リスト内のカテゴリのみです。 API を使用して、マスター カテゴリを追加、取得、および削除できます。

> [!IMPORTANT]
> アドインがカテゴリ マスター リストを管理するには、マニフェスト `Permissions` のノードをに設定する必要があります `ReadWriteMailbox`。

### <a name="add-master-categories"></a>マスター カテゴリの追加

次の例は、"Urgent! " という名前のカテゴリを追加する方法を示しています。 mailbox.masterCategories で [addAsync](/javascript/api/outlook/office.mastercategories#outlook-office-mastercategories-addasync-member(1)) を呼び出して [マスター リストに移動します](/javascript/api/outlook/office.mailbox#outlook-office-mailbox-mastercategories-member)。

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

### <a name="get-master-categories"></a>マスター カテゴリの取得

次の例は、mailbox.masterCategories で [getAsync](/javascript/api/outlook/office.mastercategories#outlook-office-mastercategories-getasync-member(1)) を呼び出してカテゴリの一覧を [取得する方法を示しています](/javascript/api/outlook/office.mailbox#outlook-office-mailbox-mastercategories-member)。

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

### <a name="remove-master-categories"></a>マスター カテゴリの削除

次の例は、"Urgent! " という名前のカテゴリを削除する方法を示しています。 [mailbox.masterCategories](/javascript/api/outlook/office.mailbox#outlook-office-mailbox-mastercategories-member) で [removeAsync](/javascript/api/outlook/office.mastercategories#outlook-office-mastercategories-removeasync-member(1)) を呼び出してマスター リストから取得します。

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

## <a name="manage-categories-on-a-message-or-appointment"></a>メッセージまたは予定のカテゴリを管理する

API を使用して、メッセージまたは予定アイテムのカテゴリを追加、取得、および削除できます。

> [!IMPORTANT]
> メッセージまたは予定に適用できるのは、メールボックスのマスター リスト内のカテゴリのみです。 詳細については、前の [セクション「マスター リストのカテゴリを管理する](#manage-categories-in-the-master-list) 」を参照してください。
>
> このOutlook on the web API を使用して、読み取りモードでメッセージのカテゴリを管理することはできません。

### <a name="add-categories-to-an-item"></a>アイテムにカテゴリを追加する

次の例は、"Urgent! " という名前のカテゴリを適用する方法を示しています。 addAsync on を呼び [出して現在のアイテムに](/javascript/api/outlook/office.categories#outlook-office-categories-addasync-member(1)) アクセスします `item.categories`。

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

### <a name="get-an-items-categories"></a>アイテムのカテゴリを取得する

次の例は、getAsync on を呼び出して、現在のアイテムに適用される [カテゴリを取得する方法を示](/javascript/api/outlook/office.categories#outlook-office-categories-getasync-member(1)) しています `item.categories`。

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

### <a name="remove-categories-from-an-item"></a>アイテムからカテゴリを削除する

次の例は、"Urgent! " という名前のカテゴリを削除する方法を示しています。 を呼び出して、現在の [アイテムから removeAsync を呼び出](/javascript/api/outlook/office.categories#outlook-office-categories-removeasync-member(1)) します `item.categories`。

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

## <a name="see-also"></a>関連項目

- [Outlookアクセス許可](understanding-outlook-add-in-permissions.md)
- [マニフェストの Permissions 要素](/javascript/api/manifest/permissions)
