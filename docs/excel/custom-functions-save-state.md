---
ms.date: 07/10/2019
description: '`OfficeRuntime.storage`を使用し、カスタム関数で状態を保存します。'
title: カスタム関数で状態を保存して共有する
localization_priority: Priority
ms.openlocfilehash: a1b70433ef0c00d507175b32fc12603ff3de1e3f
ms.sourcegitcommit: bb44c9694f88cde32ffbb642689130db44456964
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/17/2019
ms.locfileid: "35771590"
---
# <a name="save-and-share-state-in-custom-functions"></a>カスタム関数で状態を保存して共有する

`OfficeRuntime.storage`オブジェクトを使用し、カスタム関数またはアドインの作業ウィンドウに関連した状態を保存します。 ストレージはドメイン 1 つにつき 10 MB に制限されています (複数のアドインで共有される可能性があります)。 Excel on Windows では、`storage` オブジェクトはカスタム関数ランタイムの範囲内の別の場所になりますが、Excel on the web と Mac では、`storage` オブジェクトはブラウザーの `localStorage` と同じです。

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

状態管理に`storage`を使用する方法は複数あります。

- オフラインで Web リソースにアクセスできない時でも、カスタム関数を使用するための既定値を格納できます。
- Web リソースへの追加の呼び出しを回避するために使用するカスタム関数の値を保存できます。
- カスタム関数の値を保存できます。
- 作業ウィンドウの値を格納できます。

次のコード サンプルでは、`storage`に項目を格納してそれを取得する方法を示します。

```js
function storeValue(key, value) {
  return OfficeRuntime.storage.setItem(key, value).then(function (result) {
      return "Success: Item with key '" + key + "' saved to storage.";
  }, function (error) {
      return "Error: Unable to save item with key '" + key + "' to storage. " + error;
  });
}

function GetValue(key) {
  return OfficeRuntime.storage.getItem(key);
}
```

[GitHub 上の詳細なコードサンプル](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Excel-custom-functions/AsyncStorage)では、作業ウィンドウに上記の情報を受け渡す例を紹介しています。

>[!NOTE]
> `storage`オブジェクトは、現在は推奨されていないところの`AsyncStorage`と名付けられた以前のストレージ オブジェクトの代わりとなります。 現行のカスタム関数コードで`AsyncStorage`オブジェクトを使用している場合は、それを更新して`storage`オブジェクトを使用してください。

## <a name="next-steps"></a>次の手順
[カスタム関数の JSON メタデータを自動生成する](custom-functions-json-autogeneration.md)方法を学びます。 

## <a name="see-also"></a>関連項目

* [カスタム関数のメタデータ](custom-functions-json.md)
* [Excel カスタム関数のランタイム](custom-functions-runtime.md)
* [Excel カスタム関数のチュートリアル](../tutorials/excel-tutorial-create-custom-functions.md)
* [カスタム関数のデバッグ](custom-functions-debugging.md)
