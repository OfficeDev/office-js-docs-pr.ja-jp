---
title: Outlook on Mac での Outlook アドインサポートの比較
description: Outlook on Mac でのアドインサポートと他の Outlook クライアントとの比較について学習します。
ms.date: 02/08/2021
localization_priority: Normal
ms.openlocfilehash: 83cebf20cc4ead4bb50fd1a49653ac15f8501792
ms.sourcegitcommit: fefc279b85e37463413b6b0e84c880d9ed5d7ac3
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/12/2021
ms.locfileid: "50234269"
---
# <a name="compare-outlook-add-in-support-in-outlook-on-mac-with-other-outlook-clients"></a>Outlook on Mac での Outlook アドインのサポートと他の Outlook クライアントの比較

Outlook on Mac では、Outlook on the web、Windows、iOS、Android などの他のクライアントと同じ方法で Outlook アドインを作成して実行できます。クライアントごとに JavaScript をカスタマイズする必要があります。 アドインから Office JavaScript API への同じ呼び出しは、通常、次の表で説明する領域を除き、同じように動作します。

詳細については、「[Outlook 2013 プレビューでのテスト用メール アプリの展開とインストール](testing-and-tips.md)」を参照してください。

新しい UI サポートの詳細については、新しい Mac UI 上の Outlook での [アドインのサポートを参照してください](#add-in-support-in-outlook-on-new-mac-ui-preview)。

| 分野 | Outlook on the web、Windows、およびモバイル デバイス | Outlook on Mac |
|:-----|:-----|:-----|
| サポート対象バージョンの office.js および Office アドインのマニフェスト スキーマ | Office.js および スキーマ v1.1 のすべての API。 | Office.js および スキーマ v1.1 のすべての API。<br><br>**メモ**: Outlook on Mac では、ビルド 16.35.308 以降でのみ会議の保存がサポートされます。 それ以外の場合、 `saveAsync` 新規作成モードで会議から呼び出された場合、メソッドは失敗します。 回避策については、「[Office JS API を使用して Outlook for Mac で会議を下書きとして保存できない](https://support.microsoft.com/help/4505745)」を参照してください。 |
| 定期的な予定系列のインスタンス | <ul><li>定期的な系列のマスター予定または予定インスタンスのアイテム ID および他のプロパティを取得できます。</li><li>[mailbox.displayAppointmentForm](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) を使用して、定期的な系列のインスタンスまたはマスターを表示できます。</li></ul> | <ul><li>マスター予定のアイテム ID と他のプロパティを取得できますが、定期的な系列のインスタンスのアイテム ID とプロパティは取得できません。</li><li>定期的な系列のマスター予定を表示できます。アイテム ID がない場合、定期的な系列のインスタンスは表示できません。</li></ul> |
| 予定出席者の受信者の種類 | [EmailAddressDetails.recipientType](/javascript/api/outlook/office.emailaddressdetails#recipienttype) を使用して、出席者の受信者の種類を特定できます。 | `EmailAddressDetails.recipientType` は予定出席者には `undefined` を返します。 |
| クライアント アプリケーションのバージョン文字列 | [diagnostics.hostVersion](/javascript/api/outlook/office.diagnostics#hostversion)によって返されるバージョン文字列の形式は、実際のクライアントの種類によって異なります。 以下に例を示します。<ul><li>Windows 上の Outlook: `15.0.4454.1002`</li><li>Outlook on the web: `15.0.918.2`</li></ul> |Outlook on Mac で返されるバージョン文字列 `Diagnostics.hostVersion` の例: `15.0 (140325)` |
| アイテムのカスタム プロパティ | ネットワークが使用できなくなっても、アドインはキャッシュに入っているカスタム プロパティに引き続きアクセスできます。 | Outlook on Mac はカスタム プロパティをキャッシュに入れないので、ネットワークがダウンすると、アドインはカスタム プロパティにアクセスできない可能性があります。 |
| 添付ファイルの詳細 | [AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)オブジェクトのコンテンツ タイプと添付ファイル名は、クライアントの種類によって異なります。<ul><li>`AttachmentDetails.contentType` の JSON 例: `"contentType": "image/x-png"`。 </li><li>`AttachmentDetails.name` にはファイル名拡張子は含まれません。たとえば、添付ファイルが「RE: Summer activity」という件名のメッセージの場合、添付ファイル名を表す JSON オブジェクトは `"name": "RE: Summer activity"` になります。</li></ul> | <ul><li>`AttachmentDetails.contentType` の JSON 例: `"contentType" "image/png"`</li><li>`AttachmentDetails.name` には、ファイル名拡張子が必ず含まれます。メール アイテムの添付ファイルの拡張子は .eml で、予定の拡張子は .ics です。添付ファイルが「RE: Summer activity」という件名の電子メールである場合、その添付ファイル名を表す JSON オブジェクトは `"name": "RE: Summer activity.eml"` になります。<p>**注**: アドインを介するなど、ファイルがプログラムによって拡張子なしで添付される場合、`AttachmentDetails.name` にはファイル名の一部として拡張子は含まれません。</p></li></ul> |
| `dateTimeCreated` と `dateTimeModified` のプロパティでタイム ゾーンを表す文字列 |例: `Thu Mar 13 2014 14:09:11 GMT+0800 (China Standard Time)` | 例: `Thu Mar 13 2014 14:09:11 GMT+0800 (CST)` |
| `dateTimeCreated` と `dateTimeModified` の時間精度 | 次に示すコードをアドインで使用している場合、最大の精度はミリ秒単位になります:<br/>`JSON.stringify(Office.context.mailbox.item, null, 4);`| 精度は最高で秒単位となります。 |

## <a name="add-in-support-in-outlook-on-new-mac-ui-preview"></a>新しい Mac UI での Outlook でのアドインのサポート (プレビュー)

Outlook アドインは、要件セット 1.7 まで、新しい Mac UI (プレビュー) でサポートされます。 ただし、次の要件セットと機能 **はまだ** サポートされていません。

1. API 要件セット 1.8 および 1.9
1. コンテキスト アドイン
1. 送信時
1. 作成ウィンドウのポップアウト
1. 共有フォルダーのサポート
1. `saveAsync` 会議を作成する場合

バージョン 16.38.506 から利用可能な新しい Mac UI で Outlook をプレビューしてください。 試用方法の詳細については、「Outlook for Mac - Insider Fast ビルドのリリース ノート」 [を参照してください](https://support.microsoft.com/office/d6347358-5613-433e-a49e-a9a0e8e0462a)。

次のように、どの UI バージョンを使っているのか判断できます。

**現在の UI**

&nbsp;&nbsp;&nbsp;&nbsp;![Mac の現在の UI](../images/outlook-on-mac-classic.png)

**新しい UI (プレビュー)**

&nbsp;&nbsp;&nbsp;&nbsp;![Mac プレビューの新しい UI](../images/outlook-on-mac-new.png)
