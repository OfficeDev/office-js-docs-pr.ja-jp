---
title: カテゴリの取得と設定
description: メールボックスとアイテムのカテゴリを管理する方法。
ms.date: 07/07/2022
ms.localizationpriority: medium
ms.openlocfilehash: d31cb8da4cdaf4a88141a1eac927748b1399e0d9
ms.sourcegitcommit: d8ea4b761f44d3227b7f2c73e52f0d2233bf22e2
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/11/2022
ms.locfileid: "66712826"
---
# <a name="get-and-set-categories"></a>カテゴリの取得と設定

Outlook では、ユーザーはメールボックス データを整理する手段として、メッセージと予定にカテゴリを適用できます。 ユーザーは、メールボックスの色分けされたカテゴリのマスター リストを定義し、それらのカテゴリの 1 つ以上を任意のメッセージまたは予定アイテムに適用できます。 マスター リスト内の各 [カテゴリ](/javascript/api/outlook/office.categorydetails) は、ユーザーが指定する名前と [色](/javascript/api/outlook/office.mailboxenums.categorycolor) で表されます。 Office JavaScript API を使用して、メールボックスのカテゴリ マスター リストとアイテムに適用されるカテゴリを管理できます。

> [!NOTE]
> この機能のサポートは、要件セット 1.8 で導入されました。 この要件セットをサポートする [クライアントおよびプラットフォーム](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets#requirement-sets-supported-by-exchange-servers-and-outlook-clients) を参照してください。

## <a name="manage-categories-in-the-master-list"></a>マスター リストでカテゴリを管理する

メッセージまたは予定に適用できるのは、メールボックスのマスター リスト内のカテゴリのみです。 API を使用して、マスター カテゴリの追加、取得、削除を行うことができます。

> [!IMPORTANT]
> アドインでカテゴリ マスター リストを管理するには、マニフェスト`ReadWriteMailbox`内のノードを `Permissions` .

### <a name="add-master-categories"></a>マスター カテゴリを追加する

次の例は、"Urgent!" という名前のカテゴリを追加する方法を示しています。 を使用して、[mailbox.masterCategories](/javascript/api/outlook/office.mailbox#outlook-office-mailbox-mastercategories-member) で [addAsync](/javascript/api/outlook/office.mastercategories#outlook-office-mastercategories-addasync-member(1)) を呼び出します。

```js
const masterCategoriesToAdd = [
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

### <a name="get-master-categories"></a>マスター カテゴリを取得する

次の例は、[mailbox.masterCategories](/javascript/api/outlook/office.mailbox#outlook-office-mailbox-mastercategories-member) で [getAsync](/javascript/api/outlook/office.mastercategories#outlook-office-mastercategories-getasync-member(1)) を呼び出してカテゴリの一覧を取得する方法を示しています。

```js
Office.context.mailbox.masterCategories.getAsync(function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.log("Action failed with error: " + asyncResult.error.message);
    } else {
        const masterCategories = asyncResult.value;
        console.log("Master categories:");
        masterCategories.forEach(function (item) {
            console.log("-- " + JSON.stringify(item));
        });
    }
});
```

### <a name="remove-master-categories"></a>マスター カテゴリを削除する

次の例は、"Urgent!" という名前のカテゴリを削除する方法を示しています。 [mailbox.masterCategories](/javascript/api/outlook/office.mailbox#outlook-office-mailbox-mastercategories-member) で [removeAsync](/javascript/api/outlook/office.mastercategories#outlook-office-mastercategories-removeasync-member(1)) を呼び出して、マスター リストから削除します。

```js
const masterCategoriesToRemove = ["Urgent!"];

Office.context.mailbox.masterCategories.removeAsync(masterCategoriesToRemove, function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        console.log("Successfully removed categories from master list");
    } else {
        console.log("masterCategories.removeAsync call failed with error: " + asyncResult.error.message);
    }
});
```

## <a name="manage-categories-on-a-message-or-appointment"></a>メッセージまたは予定のカテゴリを管理する

API を使用して、メッセージまたは予定アイテムのカテゴリを追加、取得、削除できます。

> [!IMPORTANT]
> メッセージまたは予定に適用できるのは、メールボックスのマスター リスト内のカテゴリのみです。 詳細については、マスター [リストの前のセクション「カテゴリの管理](#manage-categories-in-the-master-list) 」を参照してください。
>
> Outlook on the webでは、API を使用して読み取りモードでメッセージのカテゴリを管理することはできません。

### <a name="add-categories-to-an-item"></a>アイテムにカテゴリを追加する

次の例は、"Urgent!" という名前のカテゴリを適用する方法を示しています。 [で addAsync](/javascript/api/outlook/office.categories#outlook-office-categories-addasync-member(1)) を呼び出して現在の項目に変換します`item.categories`。

```js
const categoriesToAdd = ["Urgent!"];

Office.context.mailbox.item.categories.addAsync(categoriesToAdd, function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        console.log("Successfully added categories");
    } else {
        console.log("categories.addAsync call failed with error: " + asyncResult.error.message);
    }
});
```

### <a name="get-an-items-categories"></a>アイテムのカテゴリを取得する

次の例では、 [getAsync](/javascript/api/outlook/office.categories#outlook-office-categories-getasync-member(1)) on を呼び出して、現在のアイテムに適用されているカテゴリを取得する方法を `item.categories`示します。

```js
Office.context.mailbox.item.categories.getAsync(function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.log("Action failed with error: " + asyncResult.error.message);
    } else {
        const categories = asyncResult.value;
        console.log("Categories:");
        categories.forEach(function (item) {
            console.log("-- " + JSON.stringify(item));
        });
    }
});
```

### <a name="remove-categories-from-an-item"></a>アイテムからカテゴリを削除する

次の例は、"Urgent!" という名前のカテゴリを削除する方法を示しています。 で [removeAsync](/javascript/api/outlook/office.categories#outlook-office-categories-removeasync-member(1)) を呼び出して、現在の `item.categories`項目から削除します。

```js
const categoriesToRemove = ["Urgent!"];

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
- [マニフェストの Permissions 要素](/javascript/api/manifest/permissions)
