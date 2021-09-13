---
title: ホスト ページからダイアログ ボックスにメッセージを渡す別の方法
description: messageChild メソッドがサポートされていない場合に使用する回避策について説明します。
ms.date: 07/08/2021
ms.localizationpriority: medium
ms.openlocfilehash: c9382be7c591176a12bbea8269ee7d371acd5233
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/12/2021
ms.locfileid: "59150000"
---
# <a name="alternative-ways-of-passing-messages-to-a-dialog-box-from-its-host-page"></a>ホスト ページからダイアログ ボックスにメッセージを渡す別の方法

親ページから子ダイアログ ボックスにデータとメッセージを渡す推奨される方法は、「Office アドインで Office ダイアログ API を使用する」の説明に従ってメソッドを `messageChild` [使用](dialog-api-in-office-add-ins.md#pass-information-to-the-dialog-box)することです。[DialogApi 1.2](../reference/requirement-sets/dialog-api-requirement-sets.md)要件セットをサポートしていないプラットフォームまたはホストでアドインが実行されている場合は、ダイアログ ボックスに情報を渡す方法が他に 2 つあります。

- `displayDialogAsync` に渡される URL にクエリ パラメーターを追加します。
- ホスト ウィンドウとダイアログ ボックスの両方にアクセス可能な場所に情報を格納します。 2 つのウィンドウは共通のセッション ストレージ [(Window.sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage)プロパティ) を共有しないが、同じドメイン *(ポート* 番号がある場合はポート番号を含む) を持つ場合は、共通のローカル Storage を [共有します](https://www.w3schools.com/html/html5_webstorage.asp)。\*

> [!NOTE]
> \* トークン処理の戦略に影響を与えるバグがあります。 Safari または Microsoft Edge ブラウザーの **Office on the web** でアドインを実行している場合、ダイアログ ボックスとタスク ウィンドウは同じローカル ストレージを共有しないため、これらの間の通信に使用できません。

## <a name="use-local-storage"></a>ローカル ストレージの使用

ローカル ストレージを使用するには、次の例のように、呼び出しの前にホスト ページ内のオブジェクトの `setItem` `window.localStorage` `displayDialogAsync` メソッドを呼び出します。

```js
localStorage.setItem("clientID", "15963ac5-314f-4d9b-b5a1-ccb2f1aea248");
```

ダイアログ ボックスのコードは、次の例のように、必要なときにアイテムを読み取ります。

```js
var clientID = localStorage.getItem("clientID");
// You can also use property syntax:
// var clientID = localStorage.clientID;
```

## <a name="use-query-parameters"></a>クエリ パラメーターの使用

次の例は、クエリ パラメーターを使用してデータを渡す方法を示します。

```js
Office.context.ui.displayDialogAsync('https://myAddinDomain/myDialog.html?clientID=15963ac5-314f-4d9b-b5a1-ccb2f1aea248');
```

この手法を使用するサンプルについては、「[PowerPoint アドインで Microsoft Graph を使用した Excel グラフの挿入](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart)」を参照してください。

ダイアログ ボックス内のコードは、URL を解析し、パラメーター値を読み取ることができます。

> [!IMPORTANT]
> Office は、`_host_info` に渡される URL に `displayDialogAsync` というクエリ パラメーターを自動的に追加します (カスタム クエリ パラメーターが存在する場合は、その後に追加されます。 ダイアログ ボックスが移動する先の後続の URL には追加されません)。 Microsoft は、将来、この値の内容を変更したり、完全に削除したりする可能性があるため、コードでこの値の内容を読み取らないでください。 同じ値がダイアログ ボックスのセッション ストレージ [(Window.sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage) プロパティ) に追加されます。 この場合も、*コードではこの値に対する読み取りも書き込みも行わないでください*。
