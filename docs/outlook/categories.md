---
title: カテゴリの取得と設定
description: '[方法] メールボックスとアイテムのカテゴリを管理する'
ms.date: 01/14/2020
localization_priority: Normal
ms.openlocfilehash: d0bb2e9f51675c263d0a3a130c64e02e7d55b764
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/17/2020
ms.locfileid: "42721024"
---
# <a name="get-and-set-categories"></a>カテゴリの取得と設定

Outlook では、ユーザーはメールボックスのデータを整理する手段として、メッセージや予定に分類項目を適用することができます。 ユーザーは、自分のメールボックスの色分けされたカテゴリのマスターリストを定義し、そのうちの1つまたは複数のカテゴリを任意のメッセージアイテムまたは予定アイテムに適用することができます。 マスターリストの各[カテゴリ](/javascript/api/outlook/office.categorydetails)は、ユーザーが指定した名前と[色](/javascript/api/outlook/office.mailboxenums.categorycolor)で表されます。 Office JavaScript API を使用して、メールボックスのカテゴリマスターリストとアイテムに適用されるカテゴリを管理できます。

> [!NOTE]
> この機能のサポートは、要件セット1.8 で導入されました。 この要件セットをサポートする [クライアントおよびプラットフォーム](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) を参照してください。

## <a name="manage-categories-in-the-master-list"></a>マスターリストでカテゴリを管理する

メールボックスのマスターリストにあるカテゴリのみが、メッセージまたは予定に適用できます。 この API を使用して、マスターカテゴリの追加、取得、および削除を行うことができます。

> [!IMPORTANT]
> このアドインでカテゴリマスターリストを管理するには、マニフェスト内の`Permissions`ノードをに`ReadWriteMailbox`設定する必要があります。

### <a name="add-master-categories"></a>マスターカテゴリを追加する

次の例は、"至急!" という名前の分類項目を追加する方法を示しています。 をマスターリストに追加するには、[メールボックス. masterCategories](/javascript/api/outlook/office.mailbox#mastercategories)で[addasync](/javascript/api/outlook/office.mastercategories#addasync-categories--options--callback-)を呼び出します。

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

### <a name="get-master-categories"></a>マスターカテゴリを取得する

次の例は、 [getAsync](/javascript/api/outlook/office.mastercategories#getasync-options--callback-)の[メールボックス. mastercategories](/javascript/api/outlook/office.mailbox#mastercategories)で、カテゴリの一覧を取得する方法を示しています。

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

### <a name="remove-master-categories"></a>マスターシェイプカテゴリを削除する

次の例は、"至急!" という名前の分類項目を削除する方法を示しています。 マスターリストから、RemoveAsync[カテゴリ](/javascript/api/outlook/office.mailbox#mastercategories)で [ [removeAsync](/javascript/api/outlook/office.mastercategories#removeasync-categories--options--callback-) ] を呼び出します。

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

## <a name="manage-categories-on-a-message-or-appointment"></a>メッセージまたは予定の分類項目を管理する

API を使用して、メッセージアイテムまたは予定アイテムの分類項目の追加、取得、削除を行うことができます。

> [!IMPORTANT]
> メールボックスのマスターリストにあるカテゴリのみが、メッセージまたは予定に適用できます。 詳細については、前のセクション「[マスターリストでカテゴリを管理](#manage-categories-in-the-master-list)する」を参照してください。
>
> Web 上の Outlook では、API を使用してメッセージのカテゴリを閲覧モードで管理することはできません。

### <a name="add-categories-to-an-item"></a>アイテムに分類項目を追加する

次の例は、"至急!" という名前の分類項目を適用する方法を示しています。 で`item.categories` [addasync](/javascript/api/outlook/office.categories#addasync-categories--options--callback-)を呼び出して、現在のアイテムに追加します。

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

次の例は、 [getAsync](/javascript/api/outlook/office.categories#getasync-options--callback-) on `item.categories`を呼び出すことによって、現在のアイテムに適用されているカテゴリを取得する方法を示しています。

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

次の例は、"至急!" という名前の分類項目を削除する方法を示しています。 現在のアイテムから[removeAsync](/javascript/api/outlook/office.categories#removeasync-categories--options--callback-) on `item.categories`を呼び出します。

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

- [Outlook のアクセス許可](understanding-outlook-add-in-permissions.md)
- [マニフェストの Permissions 要素](../reference/manifest/permissions.md)
