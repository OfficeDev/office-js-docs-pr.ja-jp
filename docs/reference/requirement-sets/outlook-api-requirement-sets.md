---
title: Outlook JavaScript API の要件セット
description: Outlook JavaScript API の要件セットの詳細。
ms.date: 05/17/2021
ms.prod: outlook
localization_priority: Priority
ms.openlocfilehash: 967bc0590d6cf1f513352c4a5c22eeb5205c2cbe
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/21/2021
ms.locfileid: "52591997"
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

すべての Outlook API は、`Mailbox`[要件セット](../../develop/specify-office-hosts-and-api-requirements.md)に属しています。`Mailbox` 要件セットには複数のバージョンがあり、リリースされる API の新しいセットはそれぞれのセットの上位バージョンに属します。すべての Outlook クライアントが最新の API のセットをサポートするわけではありませんが、Outlook クライアントが要件セットのサポートを宣言する場合は、一般的にその要件セットのすべての API がサポートされます (例外については、特定の API または機能のマニュアルを参照してください)。

マニフェストに要件セットの最小バージョンを設定することで、アドインが表示される Outlook クライアントをコントロールできます。クライアントが最小要件セットをサポートしない場合、アドインはロードされません。たとえば、要件セットのバージョン 1.3 が指定されている場合、1.3 以上をサポートしていない Outlook クライアントには表示されません。

> [!NOTE]
> 番号付きの要件セットで API を使用するには、CDN: https://appsforoffice.microsoft.com/lib/1/hosted/office.js) で **実稼働** ライブラリを参照してください。
>
> プレビューの API の使用に関する詳細については、この記事の「[プレビュー API の使用](#using-preview-apis)」セクションを参照してください。

## <a name="using-apis-from-later-requirement-sets"></a>後続の要件セットからの API の使用

要件セットを設定しても、アドインを使用できる API は制限されません。たとえば、アドインでは要件セット "メールボックス 1.1"が指定されていて、"メールボックス 1.3" をサポートしている Outlook クライアントで実行されている場合、アドインは要件セット "メールボックス 1.3" の API を使用できます。

より新しい API を使用するために、開発者は次の操作を行うことによって、特定のアプリケーションが要件セットをサポートしているかどうかをチェックできます。

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

このセクションでは、Exchange サーバーと Outlook クライアントでサポートされる一連の要件セットについて説明します。 Outlook アドインを実行するためのサーバーおよびクライアントの要件の詳細については、「[Outlook アドインの要件](../../outlook/add-in-requirements.md)」を参照してください。

> [!IMPORTANT]
> ターゲットとなる Exchange サーバーと Outlook クライアントが異なる要件セットをサポートしている場合、低い要件セットの範囲に制限されます。 たとえば、アドインが Exchange 2013 (最高要件セット: 1.1) に対して Mac 上の Outlook 2016 (最高要件セット: 1.6) で実行されている場合、アドインは要件セット 1.1 に制限されます。

### <a name="exchange-server-support"></a>Exchange server サポート

以下のサーバーは、Outlook のアドインをサポートしています。

| 製品 | Exchange のメジャー バージョン | サポートされる API の要件セット |
|---|---|---|
| Exchange Online | 最新のビルド | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md), [1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md), [1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md), [1.6](../objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md), [1.7](../objectmodel/requirement-set-1.7/outlook-requirement-set-1.7.md), [1.8](../objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md), [1.9](../objectmodel/requirement-set-1.9/outlook-requirement-set-1.9.md), [1.10](../objectmodel/requirement-set-1.10/outlook-requirement-set-1.10.md)<br>[IdentityAPI 1.3](../requirement-sets/identity-api-requirement-sets.md)\* |
| オンプレミスの Exchange | 2019 | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md)、[1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md)、[1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md)、[1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md)、[1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md) |
|| 2016 | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md)、[1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md)、[1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md)、[1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md)、[1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md) |
|| 2013 | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md) |

> [!NOTE]
> \* アドイン コードで Identity API セット 1.3 を要求するには、`isSetSupported('IdentityAPI', '1.3')` を呼び出してサポートされているかどうかを確認します。 アドイン マニフェストでの宣言はサポートされていません。 `undefined` ではないことを確認することで、API がサポートされているかどうかを判断することもできます。 詳細については、「[後続の要件セットからの API の使用](#using-apis-from-later-requirement-sets)」を参照してください。

### <a name="outlook-client-support"></a>Outlook クライアント サポート

アドインは、以下のプラットフォーム上の Outlook でサポートされています。

| プラットフォーム | Office/Outlook のメジャー バージョン | サポートされる API の要件セット |
|---|---|---|
| Windows | Microsoft 365 サブスクリプション | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md), [1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md), [1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md), [1.6](../objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md), [1.7](../objectmodel/requirement-set-1.7/outlook-requirement-set-1.7.md), [1.8](../objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md)<sup>1</sup>, [1.9](../objectmodel/requirement-set-1.9/outlook-requirement-set-1.9.md)<sup>1</sup>, [1.10](../objectmodel/requirement-set-1.10/outlook-requirement-set-1.10.md)<sup>1</sup><br>[IdentityAPI 1.3](../requirement-sets/identity-api-requirement-sets.md)<sup>2</sup> |
|| 2019 1 回限りの購入 (製品版) | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md)、[1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md)、[1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md)、[1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md)、[1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md)、[1.6](../objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md)、[1.7](../objectmodel/requirement-set-1.7/outlook-requirement-set-1.7.md)、 [1.8](../objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md)<sup>1</sup> |
|| 2019 1 回限りの購入 (ボリューム ライセンス版) | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md)、[1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md)、[1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md)、[1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md)、[1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md)、[1.6](../objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md)、[1.7](../objectmodel/requirement-set-1.7/outlook-requirement-set-1.7.md) |
|| 2016 の 1 回限りの購入 | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md)、[1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md)、[1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md)、[1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md)<sup>3</sup> |
|| 2013 年の 1 回限りの購入 | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md)、[1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md)、[1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md)<sup>3</sup>、[1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md)<sup>3</sup> |
| Mac | 現在の UI<br>(Microsoft 365 サブスクリプションに接続) | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md)、[1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md)、[1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md)、[1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md)、[1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md)、[1.6](../objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md)、[1.7](../objectmodel/requirement-set-1.7/outlook-requirement-set-1.7.md)、 [1.8](../objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md)<br>[IdentityAPI 1.3](../requirement-sets/identity-api-requirement-sets.md)<sup>2</sup> |
|| 新しい UI (プレビュー)<sup>4</sup><br>(Microsoft 365 サブスクリプションに接続) | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md)、[1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md)、[1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md)、[1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md)、[1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md)、[1.6](../objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md)、[1.7](../objectmodel/requirement-set-1.7/outlook-requirement-set-1.7.md)、 [1.8](../objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md)<br>[IdentityAPI 1.3](../requirement-sets/identity-api-requirement-sets.md)<sup>2</sup> |
|| 2019 の 1 回限りの購入 | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md)、[1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md)、[1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md)、[1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md)、[1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md)、[1.6](../objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md) |
|| 2016 の 1 回限りの購入 | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md)、[1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md)、[1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md)、[1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md)、[1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md)、[1.6](../objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md) |
| iOS | Microsoft 365 サブスクリプション | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md)、[1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md)、[1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md)、[1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md)、[1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md)<sup>5</sup> |
| Android | Microsoft 365 サブスクリプション | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md)、[1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md)、[1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md)、[1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md)、[1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md)<sup>5</sup> |
| Web ブラウザー | 接続時の最新の Outlook UI<br>Exchange Online: Microsoft 365 サブスクリプション、Outlook.com | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md), [1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md), [1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md), [1.6](../objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md), [1.7](../objectmodel/requirement-set-1.7/outlook-requirement-set-1.7.md), [1.8](../objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md), [1.9](../objectmodel/requirement-set-1.9/outlook-requirement-set-1.9.md), [1.10](../objectmodel/requirement-set-1.10/outlook-requirement-set-1.10.md)<br>[IdentityAPI 1.3](../requirement-sets/identity-api-requirement-sets.md)<sup>2</sup> |
|| 接続時の従来の Outlook UI<br>オンプレミスの Exchange | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md)、[1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md)、[1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md)、[1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md)、[1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md)、[1.6](../objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md) |

> [!NOTE]
> <sup>1</sup> Microsoft 365 サブスクリプションまたは製品版の 1 回限りの購入による Outlook on Windows での **1.8** のサポートは、バージョン 1910 (ビルド 12130.20272) から利用できます。 Microsoft 365 サブスクリプションを使用した Windows 用 Outlook での **1.9** のサポートは、バージョン 2008 (ビルド 13127.20296) から入手できます。 Microsoft 365 サブスクリプションを使用した Windows 用 Outlook での **1.10** のサポートは、バージョン 2104 (ビルド 13929.20296) から入手できます。 バージョンに応じた詳細については、[Office 2019](/officeupdates/update-history-office-2019) または [Microsoft 365](/officeupdates/update-history-office365-proplus-by-date) の更新履歴ページと、[Office クライアントのバージョンを見つけてチャネルを更新する方法](https://support.microsoft.com/office/932788b8-a3ce-44bf-bb09-e334518b8b19)を参照してください。
>
> <sup>2</sup> アドイン コードで Identity API セット 1.3 を要求するには、`isSetSupported('IdentityAPI', '1.3')` を呼び出してサポートされているかどうかを確認します。 アドイン マニフェストでの宣言はサポートされていません。 `undefined` ではないことを確認することで、API がサポートされているかどうかを判断することもできます。 詳細については、「[後続の要件セットからの API の使用](#using-apis-from-later-requirement-sets)」を参照してください。
>
> <sup>3</sup> Outlook 2013 での 1.3 のサポートは、「[2015 年 12 月 8 日付、Outlook 2013 用更新プログラム (KB3114349)](https://support.microsoft.com/kb/3114349)」の一部として追加されました。 Outlook 2013 での 1.4 のサポートは、「[MS16-107: Outlook 2013 セキュリティ更新プログラムについて 2016 年 9 月 13 日](https://support.microsoft.com/help/3118280)」の一部として追加されました。 Outlook 2016 (1 回限りの購入) での 1.4 のサポートは、「[2018 年 7 月 3 日更新プログラム Office 2016 (KB4022223)](https://support.microsoft.com/help/4022223)」の一部として追加されました。
>
> <sup>4</sup> 新しい Mac 版 Outlook のプレビュー サポートは、バージョン 16.38.506 から利用できます。 詳細については、「[新しい Mac UI での Outlook のアドインのサポート](../../outlook/compare-outlook-add-in-support-in-outlook-for-mac.md#add-in-support-in-outlook-on-new-mac-ui-preview)」セクションを参照してください。
>
> <sup>5</sup> 現在、モバイル クライアント用のアドインを設計および実装する際には、さらに考慮事項があります。 たとえば、サポートされるモードは、メールの読み取りのみです。 詳細については、[Outlook Mobile にアドイン コマンドのサポートを追加するときのコードの考慮事項](../../outlook/add-mobile-support.md#code-considerations)を参照してください。

> [!TIP]
> メールボックスのツールバーを確認することで、Web ブラウザーでの Outlook がモダンかクラシックかを区別できます。
>
> **モダン**
>
> ![Outlook ツールバー (モダン) の部分的なスクリーンショット](../../images/outlook-on-the-web-new-toolbar.png)
>
> **クラシック**
>
> ![Outlook ツールバー (クラシック) の部分的なスクリーンショット](../../images/outlook-on-the-web-classic-toolbar.png)

## <a name="using-preview-apis"></a>プレビュー API の使用

新しい Outlook JavaScript API は最初に "プレビュー" で導入され、その後、十分なテストが行われ、ユーザー フィードバックが得られてから、番号付きの特定の要件セットの一部になります。 プレビュー API についてフィードバックを提供するには、その API が記載されている Web ページの最後にあるフィードバック メカニズムを使用してください。

> [!NOTE]
> プレビュー API は変更されることがあります。運用環境での使用は意図されていません。

プレビュー API の詳細については、「[Outlook API 要件セットのプレビュー](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md)」を参照してください。
