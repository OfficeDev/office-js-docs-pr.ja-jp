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
# <a name="alternative-ways-of-passing-messages-to-a-dialog-box-from-its-host-page"></a>ホストページからダイアログボックスにメッセージを渡す代替方法

親ページから子ダイアログボックスにデータおよびメッセージを渡す方法としては、 `messageChild` 「 [office アドインで OFFICE ダイアログ API を使用](dialog-api-in-office-add-ins.md#pass-information-to-the-dialog-box)する」で説明されている方法を使用することをお勧めします。Add-in [api 1.2 要件セット](../reference/requirement-sets/dialog-api-requirement-sets.md)をサポートしていないプラットフォームまたはホストでアドインが実行されている場合は、次の2つの方法で情報をダイアログボックスに渡すことができます。

- `displayDialogAsync` に渡される URL にクエリ パラメーターを追加します。
- ホスト ウィンドウとダイアログ ボックスの両方にアクセス可能な場所に情報を格納します。 2つのウィンドウは、共通のセッション記憶域 ( [Window. sessionstorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage) プロパティ) を共有しませんが、 *同じドメイン* (ポート番号を含む) がある場合は、共通の [ローカル記憶域](https://www.w3schools.com/html/html5_webstorage.asp)を共有します。\*


> [!NOTE]
> \* トークン処理の戦略に影響を与えるバグがあります。 Safari または Microsoft Edge ブラウザーの **Office on the web** でアドインを実行している場合、ダイアログ ボックスとタスク ウィンドウは同じローカル ストレージを共有しないため、これらの間の通信に使用できません。

## <a name="use-local-storage"></a>ローカル ストレージの使用

ローカルストレージを使用するには、 `setItem` 次の `window.localStorage` `displayDialogAsync` 例に示すように、呼び出しの前に、ホストページでオブジェクトのメソッドを呼び出します。

```js
localStorage.setItem("clientID", "15963ac5-314f-4d9b-b5a1-ccb2f1aea248");
```

ダイアログ ボックス内のコードは、次の例に示すように、必要に応じて項目を読み取ります。

```js
var clientID = localStorage.getItem("clientID");
// You can also use property syntax:
// var clientID = localStorage.clientID;
```

## <a name="use-query-parameters"></a>クエリ パラメーターの使用

次の例では、クエリ パラメーターを使用してデータを渡す方法を示します。

```js
Office.context.ui.displayDialogAsync('https://myAddinDomain/myDialog.html?clientID=15963ac5-314f-4d9b-b5a1-ccb2f1aea248');
```

この手法を使用するサンプルについては、「[PowerPoint アドインで Microsoft Graph を使用した Excel グラフの挿入](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart)」を参照してください。

ダイアログ ボックス内のコードは、URL を解析し、パラメーター値を読み取ることができます。

> [!IMPORTANT]
> Office は、`_host_info` に渡される URL に `displayDialogAsync` というクエリ パラメーターを自動的に追加します (カスタム クエリ パラメーターが存在する場合は、その後に追加されます。 ダイアログ ボックスが移動する先の後続の URL には追加されません)。 Microsoft は、将来、この値の内容を変更したり、完全に削除したりする可能性があるため、コードでこの値の内容を読み取らないでください。 ダイアログボックスのセッションストレージ ( [sessionstorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage) プロパティ) に同じ値が追加されます。 この場合も、*コードではこの値に対する読み取りも書き込みも行わないでください*。
