---
title: Outlook on Mac での Outlook アドインのサポートを比較する
description: Outlook on Mac のアドイン サポートと他の Outlook クライアントの比較方法について説明します。
ms.date: 09/21/2022
ms.localizationpriority: medium
ms.openlocfilehash: c3f991865921583561e4c2db2132fad3ceba3625
ms.sourcegitcommit: 09bb0b5edd6af03c9822e1742095c7df94735120
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/23/2022
ms.locfileid: "67990414"
---
# <a name="compare-outlook-add-in-support-in-outlook-on-mac-with-other-outlook-clients"></a>Outlook on Mac での Outlook アドインのサポートと他の Outlook クライアントの比較

Outlook アドインは、クライアントごとに JavaScript をカスタマイズせずに、他のクライアント (Outlook on the web、Windows、iOS、Android など) と同じように Outlook on Mac で作成および実行できます。 アドインから Office JavaScript API への同じ呼び出しは、通常、次の表で説明する領域を除き、同じように動作します。

詳細については、「[Outlook 2013 プレビューでのテスト用メール アプリの展開とインストール](testing-and-tips.md)」を参照してください。

新しい UI サポートの詳細については、 [Outlook on new Mac UI のアドイン サポートに関するページを](#add-in-support-in-outlook-on-new-mac-ui)参照してください。

| 分野 | Outlook on the web、Windows、モバイル デバイス | Outlook on Mac |
|:-----|:-----|:-----|
| サポート対象バージョンの office.js および Office アドインのマニフェスト スキーマ | Office.js および スキーマ v1.1 のすべての API。 | Office.js および スキーマ v1.1 のすべての API。<br><br>**注**: Outlook on Mac では、ビルド 16.35.308 以降でのみ会議の保存がサポートされます。 それ以外の `saveAsync` 場合、作成モードで会議から呼び出されると、メソッドは失敗します。 回避策については、「[Office JS API を使用して Outlook for Mac で会議を下書きとして保存できない](https://support.microsoft.com/help/4505745)」を参照してください。 |
| 定期的な予定系列のインスタンス | <ul><li>定期的な系列のマスター予定または予定インスタンスのアイテム ID および他のプロパティを取得できます。</li><li>[mailbox.displayAppointmentForm](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) を使用して、定期的な系列のインスタンスまたはマスターを表示できます。</li></ul> | <ul><li>マスター予定のアイテム ID と他のプロパティを取得できますが、定期的な系列のインスタンスのアイテム ID とプロパティは取得できません。</li><li>Can display the master appointment of a recurring series. Without the item ID, cannot display an instance of a recurring series.</li></ul> |
| 予定出席者の受信者の種類 | [EmailAddressDetails.recipientType](/javascript/api/outlook/office.emailaddressdetails#outlook-office-emailaddressdetails-recipienttype-member) を使用して、出席者の受信者の種類を特定できます。 | `EmailAddressDetails.recipientType` は予定出席者には `undefined` を返します。 |
| クライアント アプリケーションのバージョン文字列 | [diagnostics.hostVersion](/javascript/api/outlook/office.diagnostics#outlook-office-diagnostics-hostversion-member) によって返されるバージョン文字列の形式は、クライアントの実際の種類によって異なります。 以下に例を示します。<ul><li>Outlook on Windows: `15.0.4454.1002`</li><li>Outlook on the web:`15.0.918.2`</li></ul> |Outlook on Mac で返される `Diagnostics.hostVersion` バージョン文字列の例を次に示します。 `15.0 (140325)` |
| アイテムのカスタム プロパティ | ネットワークが使用できなくなっても、アドインはキャッシュに入っているカスタム プロパティに引き続きアクセスできます。 | Outlook on Mac ではカスタム プロパティがキャッシュされないため、ネットワークがダウンした場合、アドインはそれらにアクセスできません。 |
| 添付ファイルの詳細 | [AttachmentDetails](/javascript/api/outlook/office.attachmentdetails) オブジェクトのコンテンツ タイプと添付ファイル名は、クライアントの種類によって異なります。<ul><li>`AttachmentDetails.contentType` の JSON 例: `"contentType": "image/x-png"`。 </li><li>`AttachmentDetails.name` does not contain any filename extension. As an example, if the attachment is a message that has the subject "RE: Summer activity", the JSON object that represents the attachment name would be `"name": "RE: Summer activity"`.</li></ul> | <ul><li>`AttachmentDetails.contentType` の JSON 例: `"contentType" "image/png"`</li><li>`AttachmentDetails.name` always includes a filename extension. Attachments that are mail items have a .eml extension, and appointments have a .ics extension. As an example, if an attachment is an email with the subject "RE: Summer activity", the JSON object that represents the attachment name would be `"name": "RE: Summer activity.eml"`.<p>**注**: アドインを介するなど、ファイルがプログラムによって拡張子なしで添付される場合、`AttachmentDetails.name` にはファイル名の一部として拡張子は含まれません。</p></li></ul> |
| `dateTimeCreated` と `dateTimeModified` のプロパティでタイム ゾーンを表す文字列 |例: `Thu Mar 13 2014 14:09:11 GMT+0800 (China Standard Time)` | 例: `Thu Mar 13 2014 14:09:11 GMT+0800 (CST)` |
| `dateTimeCreated` と `dateTimeModified` の時間精度 | 次に示すコードをアドインで使用している場合、最大の精度はミリ秒単位になります。<br/>`JSON.stringify(Office.context.mailbox.item, null, 4);`| 精度は最大で秒単位となります。 |

## <a name="add-in-support-in-outlook-on-new-mac-ui"></a>Outlook on new Mac UI でのアドインのサポート

Outlook アドインは、新しい Mac UI でサポートされるようになりました (Outlook バージョン 16.38.506 から入手できます)。 新しい Mac UI で現在サポートされている要件セットについては、 [Outlook API 要件セットクライアントのサポート](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets#outlook-client-support)に関するページを参照してください。

新しい Mac UI の詳細については、「[新しいOutlook for Mac](https://support.microsoft.com/office/6283be54-e74d-434e-babb-b70cefc77439)」を参照してください。

次のように、どの UI バージョンを使用しているかを確認できます。

**クラシック UI**

![Mac 上のクラシック UI。](../images/outlook-on-mac-classic.png)

**新しい UI**

![Mac の新しい UI。](../images/outlook-on-mac-new.png)
