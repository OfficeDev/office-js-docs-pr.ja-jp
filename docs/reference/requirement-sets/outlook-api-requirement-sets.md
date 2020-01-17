---
title: Outlook JavaScript API の要件セット
description: ''
ms.date: 01/14/2020
ms.prod: outlook
localization_priority: Priority
ms.openlocfilehash: bd6b20e9f0ddb5141f2f889a4e99af2c042a10ab
ms.sourcegitcommit: 212c810f3480a750df779777c570159a7f76054a
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/17/2020
ms.locfileid: "41217373"
---
# <a name="outlook-javascript-api-requirement-sets"></a>Outlook JavaScript API の要件セット

Outlook アドインは、[マニフェスト](../manifest/requirements.md)で[要件](../../develop/add-in-manifests.md)要素を使用して、必要な API のバージョンを宣言します。Outlook アドインには、`Name` 属性が `Mailbox` に設定され、`MinVersion` 属性がアドインのシナリオをサポートする最小 API 要件セットに設定された[設定](../manifest/set.md)要素が常に含まれます。

たとえば、次のマニフェストのスニペットは、最小要件セットの 1.1 を表します。

```xml
<Requirements>
  <Sets>
    <Set Name="Mailbox" MinVersion="1.1" />
  </Sets>
</Requirements>
```

すべての Outlook API は、`Mailbox`[要件セット](../../develop/specify-office-hosts-and-api-requirements.md)に属しています。`Mailbox` 要件セットには複数のバージョンがあり、リリースされる API の新しいセットはそれぞれのセットの上位バージョンに属します。すべての Outlook クライアントが最新の API のセットをサポートするわけではありませんが、Outlook クライアントが要件セットのサポートを宣言する場合は、その要件セットのすべての API がサポートされます。

マニフェストに要件セットの最小バージョンを設定することで、アドインが表示される Outlook クライアントをコントロールできます。クライアントが最小要件セットをサポートしない場合、アドインはロードされません。たとえば、要件セットのバージョン 1.3 が指定されている場合、1.3 以上をサポートしていない Outlook クライアントには表示されません。

> [!NOTE]
> 番号付きの要件セットで API を使用するには、CDN: https://appsforoffice.microsoft.com/lib/1/hosted/office.js) で**実稼働**ライブラリを参照してください。
>
> プレビューの API の使用に関する詳細については、この記事の「[プレビュー API の使用](#using-preview-apis)」セクションを参照してください。

## <a name="using-apis-from-later-requirement-sets"></a>後続の要件セットからの API の使用

要件セットを設定しても、アドインで使用できる API は制限されません。 たとえば、アドインでは要件セット「Mailbox 1.1」が指定されていて、「Mailbox 1.3」をサポートしている Outlook クライアントで実行されている場合、アドインは要件セット「Mailbox 1.3」の API を使用できます。

より新しい API を使用するために、開発者は次の操作を行うことによって、特定のホストが要件セットをサポートしているかどうかをチェックできます。

```js
if (Office.context.requirements.isSetSupported('Mailbox', '1.3')) {
  // Perform actions.
}
else {
  // Provide alternate flow/logic.
}
```

または、開発者は標準の JavaScript の技法を使用することで、新しい API の有無をチェックできます。

```js
if (item.somePropertyOrFunction !== undefined) {
  // Use item.somePropertyOrFunction.
  item.somePropertyOrFunction;
}
```

このようなチェックは、マニフェストで指定された要件セット バージョンに存在する API には必要ありません。

## <a name="choosing-a-minimum-requirement-set"></a>最小要件セットの選択

開発者は、アドインを使用するために必要な、シナリオで必須の API のセットが含まれている初期の要件セットを使用する必要があります。

## <a name="requirement-sets-supported-by-exchange-servers-and-outlook-clients"></a>Exchange サーバーと Outlook クライアントでサポートされる要件セット

このセクションでは、Exchange サーバーと Outlook クライアントでサポートされる一連の要件セットについて説明します。

> [!IMPORTANT]
> ターゲットとなる Exchange サーバーと Outlook クライアントが異なる要件セットをサポートしている場合、低い要件セットの範囲に制限されます。 たとえば、アドインが Exchange 2013 (最高要件セット: 1.1) に対して Mac 上の Outlook 2016 (最高要件セット: 1.6) で実行されている場合、アドインは要件セット 1.1 に制限されます。

### <a name="exchange-server-support"></a>Exchange server サポート

以下のサーバーは、Outlook のアドインをサポートしています。

| 製品 | Exchange のメジャー バージョン | サポートされる API の要件セット |
|---|---|---|
| Exchange Online | 最新のビルド | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md)、[1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md)、[1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md)、[1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md)、[1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md)、[1.6](../objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md)、[1.7](../objectmodel/requirement-set-1.7/outlook-requirement-set-1.7.md)、 [1.8](../objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md) |
| オンプレミスの Exchange | 2019 | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md)、[1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md)、[1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md)、[1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md)、[1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md) |
|| 2016 | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md)、[1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md)、[1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md)、[1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md)、[1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md) |
|| 2013 | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="outlook-client-support"></a>Outlook クライアント サポート

アドインは、以下のプラットフォーム上の Outlook でサポートされています。

| プラットフォーム | Office/Outlook のメジャー バージョン | サブスクリプションまたは 1 回限りの購入 ? | サポートされる API の要件セット |
|---|---|---|---|
| Windows | 最新のビルド<br>(月次チャネル) | Office 365 サブスクリプション | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md)、[1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md)、[1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md)、[1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md)、[1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md)、[1.6](../objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md)、[1.7](../objectmodel/requirement-set-1.7/outlook-requirement-set-1.7.md)、 [1.8](../objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md) |
|| 2019 | 1 回限りの購入 | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md)、[1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md)、[1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md)、[1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md)、[1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md)、[1.6](../objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md)、[1.7](../objectmodel/requirement-set-1.7/outlook-requirement-set-1.7.md) |
|| 2016 | 1 回限りの購入 | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md)、[1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md)、[1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md)、[1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md) |
|| 2013 | 1 回限りの購入 | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md)、[1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md)、[1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md)、[1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md) |
| Mac | 最新のビルド<br>(月次チャネル) | Office 365 サブスクリプション | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md)、[1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md)、[1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md)、[1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md)、[1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md)、[1.6](../objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md)、[1.7](../objectmodel/requirement-set-1.7/outlook-requirement-set-1.7.md)、 [1.8](../objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md) |
|| 2019 | 1 回限りの購入 | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md)、[1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md)、[1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md)、[1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md)、[1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md)、[1.6](../objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md) |
|| 2016 | 1 回限りの購入 | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md)、[1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md)、[1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md)、[1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md)、[1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md)、[1.6](../objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md) |
| iOS | 最新のビルド<br>(月次チャネル) | Office 365 サブスクリプション | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md)、[1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md)、[1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md)、[1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md)、[1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md) |
| Android | 最新のビルド<br>(月次チャネル) | Office 365 サブスクリプション | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md)、[1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md)、[1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md)、[1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md)、[1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md) |
| Web ブラウザー | 最新 | Exchange Online: Office 365 サブスクリプション、Outlook.com | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md)、[1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md)、[1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md)、[1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md)、[1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md)、[1.6](../objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md)、[1.7](../objectmodel/requirement-set-1.7/outlook-requirement-set-1.7.md)、 [1.8](../objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md) |
|| クラシック | オンプレミスの Exchange | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md)、[1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md)、[1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md)、[1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md)、[1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md)、[1.6](../objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md) |

> [!NOTE]
> Outlook 2013 での 1.3 のサポートは、「[2015 年 12 月 8 日付、Outlook 2013 用更新プログラム (KB3114349)](https://support.microsoft.com/kb/3114349)」の一部として追加されました。 Outlook 2013 での 1.4 のサポートは、「[MS16-107: Outlook 2013 セキュリティ更新プログラムについて 2016 年 9 月 13 日](https://support.microsoft.com/help/3118280)」の一部として追加されました。 Outlook 2016 (MSI) での 1.4 のサポートは、「[2018 年 7 月 3 日更新プログラム Office 2016 (KB4022223)](https://support.microsoft.com/help/4022223)」の一部として追加されました。

> [!TIP]
> メールボックスのツールバーを確認することで、Web ブラウザーでの Outlook がモダンかクラシックかを区別できます。
>
> **モダン**
>
> ![Outlook ツールバー (モダン) の部分的なスクリーンショット](https://docs.microsoft.com/outlook/add-ins/images/outlook-on-the-web-new-toolbar.png)
>
> **クラシック**
>
> ![Outlook ツールバー (クラシック) の部分的なスクリーンショット](https://docs.microsoft.com/outlook/add-ins/images/outlook-on-the-web-classic-toolbar.png)

## <a name="using-preview-apis"></a>プレビュー API の使用

新しい Outlook JavaScript API は最初に "プレビュー" で導入され、その後、十分なテストが行われ、ユーザー フィードバックが得られてから、番号付きの特定の要件セットの一部になります。 プレビュー API についてフィードバックを提供するには、その API が記載されている Web ページの最後にあるフィードバック メカニズムを使用してください。

> [!NOTE]
> プレビュー API は変更されることがあります。運用環境での使用は意図されていません。

プレビュー API の詳細については、「[Outlook API 要件セットのプレビュー](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md)」を参照してください。
