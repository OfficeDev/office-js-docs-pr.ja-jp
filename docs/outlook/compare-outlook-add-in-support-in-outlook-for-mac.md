---
title: Mac 上Outlook Outlookでのアドインのサポートを比較する
description: Mac 上のOutlookでのアドインのサポートが他のOutlook クライアントと比較する方法について説明します。
ms.date: 06/09/2022
ms.localizationpriority: medium
ms.openlocfilehash: 36a10f0454bebf3f069464277c7eb2a8a18f42b7
ms.sourcegitcommit: 2eeb0423a793b3a6db8a665d9ae6bcb10e867be3
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/10/2022
ms.locfileid: "66019606"
---
# <a name="compare-outlook-add-in-support-in-outlook-on-mac-with-other-outlook-clients"></a>Mac 上の Outlook でのアドイン サポートOutlook他のOutlook クライアントと比較する

Outlook アドインは、他のクライアント (Outlook on the web、Windows、iOS、Androidなど) と同じ方法で Mac 上のOutlookで作成および実行できます。クライアントごとに JavaScript をカスタマイズする必要はありません。 アドインから Office JavaScript API への同じ呼び出しは、通常、次の表で説明する領域を除き、同じように動作します。

詳細については、「[Outlook 2013 プレビューでのテスト用メール アプリの展開とインストール](testing-and-tips.md)」を参照してください。

新しい UI サポートの詳細については、新[しい Mac UI のOutlookでのアドインサポートに関するページを](#add-in-support-in-outlook-on-new-mac-ui)参照してください。

| 分野 | Outlook on the web、Windows、モバイル デバイス | Outlook on Mac |
|:-----|:-----|:-----|
| サポート対象バージョンの office.js および Office アドインのマニフェスト スキーマ | Office.js および スキーマ v1.1 のすべての API。 | Office.js および スキーマ v1.1 のすべての API。<br><br>**注**: Mac のOutlookでは、ビルド 16.35.308 以降でのみ会議の保存がサポートされます。 それ以外の `saveAsync` 場合、作成モードで会議から呼び出されると、メソッドは失敗します。 回避策については、「[Office JS API を使用して Outlook for Mac で会議を下書きとして保存できない](https://support.microsoft.com/help/4505745)」を参照してください。 |
| 定期的な予定系列のインスタンス | <ul><li>定期的な系列のマスター予定または予定インスタンスのアイテム ID および他のプロパティを取得できます。</li><li>[mailbox.displayAppointmentForm](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) を使用して、定期的な系列のインスタンスまたはマスターを表示できます。</li></ul> | <ul><li>マスター予定のアイテム ID と他のプロパティを取得できますが、定期的な系列のインスタンスのアイテム ID とプロパティは取得できません。</li><li>定期的な系列のマスター予定を表示できます。アイテム ID がない場合、定期的な系列のインスタンスは表示できません。</li></ul> |
| 予定出席者の受信者の種類 | [EmailAddressDetails.recipientType](/javascript/api/outlook/office.emailaddressdetails#outlook-office-emailaddressdetails-recipienttype-member) を使用して、出席者の受信者の種類を特定できます。 | `EmailAddressDetails.recipientType` は予定出席者には `undefined` を返します。 |
| クライアント アプリケーションのバージョン文字列 | [diagnostics.hostVersion](/javascript/api/outlook/office.diagnostics#outlook-office-diagnostics-hostversion-member) によって返されるバージョン文字列の形式は、クライアントの実際の種類によって異なります。 次に例を示します。<ul><li>WindowsのOutlook:`15.0.4454.1002`</li><li>Outlook on the web:`15.0.918.2`</li></ul> |Mac のOutlookで`Diagnostics.hostVersion`返されるバージョン文字列の例を次に示します。`15.0 (140325)` |
| アイテムのカスタム プロパティ | ネットワークが使用できなくなっても、アドインはキャッシュに入っているカスタム プロパティに引き続きアクセスできます。 | Mac 上のOutlookはカスタム プロパティをキャッシュしないため、ネットワークがダウンした場合、アドインはそれらにアクセスできなくなります。 |
| 添付ファイルの詳細 | [AttachmentDetails](/javascript/api/outlook/office.attachmentdetails) オブジェクトのコンテンツ タイプと添付ファイル名は、クライアントの種類によって異なります。<ul><li>`AttachmentDetails.contentType` の JSON 例: `"contentType": "image/x-png"`。 </li><li>`AttachmentDetails.name` にはファイル名拡張子は含まれません。たとえば、添付ファイルが「RE: Summer activity」という件名のメッセージの場合、添付ファイル名を表す JSON オブジェクトは `"name": "RE: Summer activity"` になります。</li></ul> | <ul><li>`AttachmentDetails.contentType` の JSON 例: `"contentType" "image/png"`</li><li>`AttachmentDetails.name` には、ファイル名拡張子が必ず含まれます。メール アイテムの添付ファイルの拡張子は .eml で、予定の拡張子は .ics です。添付ファイルが「RE: Summer activity」という件名の電子メールである場合、その添付ファイル名を表す JSON オブジェクトは `"name": "RE: Summer activity.eml"` になります。<p>**注**: アドインを介するなど、ファイルがプログラムによって拡張子なしで添付される場合、`AttachmentDetails.name` にはファイル名の一部として拡張子は含まれません。</p></li></ul> |
| `dateTimeCreated` と `dateTimeModified` のプロパティでタイム ゾーンを表す文字列 |例: `Thu Mar 13 2014 14:09:11 GMT+0800 (China Standard Time)` | 例: `Thu Mar 13 2014 14:09:11 GMT+0800 (CST)` |
| `dateTimeCreated` と `dateTimeModified` の時間精度 | 次に示すコードをアドインで使用している場合、最大の精度はミリ秒単位になります。<br/>`JSON.stringify(Office.context.mailbox.item, null, 4);`| 精度は最大で秒単位となります。 |

## <a name="add-in-support-in-outlook-on-new-mac-ui"></a>新しい Mac UI でのOutlookでのアドインのサポート

Outlook アドインは、新しい Mac UI (Outlook バージョン 16.38.506 から利用可能) で、要件セット 1.10 までサポートされるようになりました。 ただし、次の要件セットと機能はまだサポート **されていません** 。

- API 要件セット 1.11

新しい Mac UI の詳細については、「[新しいOutlook for Mac](https://support.microsoft.com/office/6283be54-e74d-434e-babb-b70cefc77439)」を参照してください。

次のように、どの UI バージョンを使用しているかを確認できます。

**クラシック UI**

![Mac 上のクラシック UI。](../images/outlook-on-mac-classic.png)

**新しい UI**

![Mac の新しい UI。](../images/outlook-on-mac-new.png)
