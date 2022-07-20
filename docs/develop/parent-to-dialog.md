---
title: ホスト ページからダイアログ ボックスにメッセージを渡す別の方法
description: messageChild メソッドがサポートされていない場合に使用する回避策について説明します。
ms.date: 07/18/2022
ms.localizationpriority: medium
ms.openlocfilehash: f42a549a3c39866516cfd5395dd7589a890b0956
ms.sourcegitcommit: df7964b6509ee6a807d754fbe895d160bc52c2d3
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/20/2022
ms.locfileid: "66889416"
---
# <a name="alternative-ways-of-passing-messages-to-a-dialog-box-from-its-host-page"></a>ホスト ページからダイアログ ボックスにメッセージを渡す別の方法

親ページから子ダイアログ ボックスにデータとメッセージを渡す場合に推奨される方法は `messageChild` 、「 [Office アドインで Office ダイアログ API を使用する」の](dialog-api-in-office-add-ins.md#pass-information-to-the-dialog-box)説明に従ってメソッドを使用することです。 [アドインが DialogApi 1.2 要件セット](/javascript/api/requirement-sets/common/dialog-api-requirement-sets)をサポートしていないプラットフォームまたはホストで実行されている場合は、他に 2 つの方法で情報をダイアログ ボックスに渡すことができます。

- `displayDialogAsync` に渡される URL にクエリ パラメーターを追加します。
- ホスト ウィンドウとダイアログ ボックスの両方にアクセス可能な場所に情報を格納します。 2 つのウィンドウは共通のセッション ストレージ ( [Window.sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage) プロパティ) を共有しませんが、 *同じドメイン* (ポート番号がある場合はポート番号を含む) がある場合は、共通 [のローカル ストレージ](https://www.w3schools.com/html/html5_webstorage.asp)を共有します。\*

> [!NOTE]
> \* トークン処理の戦略に影響を与えるバグがあります。 Safari または Microsoft Edge ブラウザーの **Office on the web** でアドインを実行している場合、ダイアログ ボックスとタスク ウィンドウは同じローカル ストレージを共有しないため、これらの間の通信に使用できません。

## <a name="use-local-storage"></a>ローカル ストレージの使用

ローカル ストレージを使用するには、次の`window.localStorage`例のように、呼び出し`setItem`の前に`displayDialogAsync`ホスト ページでオブジェクトのメソッドを呼び出します。

```js
localStorage.setItem("clientID", "15963ac5-314f-4d9b-b5a1-ccb2f1aea248");
```

ダイアログ ボックスのコードは、次の例のように、必要なときにアイテムを読み取ります。

```js
const clientID = localStorage.getItem("clientID");
// You can also use property syntax:
// const clientID = localStorage.clientID;
```

## <a name="use-query-parameters"></a>クエリ パラメーターの使用

次の例は、クエリ パラメーターを使用してデータを渡す方法を示します。

```js
Office.context.ui.displayDialogAsync('https://myAddinDomain/myDialog.html?clientID=15963ac5-314f-4d9b-b5a1-ccb2f1aea248');
```

この手法を使用するサンプルについては、「[PowerPoint アドインで Microsoft Graph を使用した Excel グラフの挿入](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart)」を参照してください。

ダイアログ ボックス内のコードは、URL を解析し、パラメーター値を読み取ることができます。

> [!IMPORTANT]
> Office は、`_host_info` に渡される URL に `displayDialogAsync` というクエリ パラメーターを自動的に追加します (カスタム クエリ パラメーターが存在する場合は、その後に追加されます。 ダイアログ ボックスが移動する先の後続の URL には追加されません)。 Microsoft は、将来、この値の内容を変更したり、完全に削除したりする可能性があるため、コードでこの値の内容を読み取らないでください。 同じ値がダイアログ ボックスのセッション ストレージ ( [Window.sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage) プロパティ) に追加されます。 この場合も、*コードではこの値に対する読み取りも書き込みも行わないでください*。
