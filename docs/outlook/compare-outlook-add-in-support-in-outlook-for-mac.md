---
title: Outlook on Mac での Outlook アドインサポートの比較
description: Outlook on Mac でのアドインのサポートが他の Outlook ホストと比較する方法について説明します。
ms.date: 11/26/2019
localization_priority: Normal
ms.openlocfilehash: 452a162a05a44477c85a99417b3a8cefbf49e515
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720828"
---
# <a name="compare-outlook-add-in-support-in-outlook-on-mac-with-other-outlook-hosts"></a>Outlook on Mac での outlook アドインのサポートと他の Outlook ホストの比較

Outlook on the web、Windows、iOS、Android などの他のホストと同じ方法で Outlook アドインを作成して実行すると、各ホストの JavaScript をカスタマイズする必要はありません。 通常、アドインから Office JavaScript API への同じ呼び出しは、次の表に示す領域を除き、同じように動作します。

詳細については、「[Outlook 2013 プレビューでのテスト用メール アプリの展開とインストール](testing-and-tips.md)」を参照してください。

| 分野 | Web 上の Outlook、Windows、およびモバイルデバイス | Outlook on Mac |
|:-----|:-----|:-----|
| サポート対象バージョンの office.js および Office アドインのマニフェスト スキーマ | Office.js および スキーマ v1.1 のすべての API。 | Office.js および スキーマ v1.1 のすべての API。<br><br>**注**: Outlook on Mac では、会議の保存はサポートされていません。 `saveAsync` メソッドは、作成モードの会議から呼び出されると失敗します。 回避策については、「[Office JS API を使用して Outlook for Mac で会議を下書きとして保存できない](https://support.microsoft.com/help/4505745)」を参照してください。 |
| 定期的な予定系列のインスタンス | <ul><li>定期的な系列のマスター予定または予定インスタンスのアイテム ID および他のプロパティを取得できます。</li><li>[mailbox.displayAppointmentForm](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) を使用して、定期的な系列のインスタンスまたはマスターを表示できます。</li></ul> | <ul><li>マスター予定のアイテム ID と他のプロパティを取得できますが、定期的な系列のインスタンスのアイテム ID とプロパティは取得できません。</li><li>定期的な系列のマスター予定を表示できます。アイテム ID がない場合、定期的な系列のインスタンスは表示できません。</li></ul> |
| 予定出席者の受信者の種類 | [EmailAddressDetails.recipientType](/javascript/api/outlook/office.emailaddressdetails#recipienttype) を使用して、出席者の受信者の種類を特定できます。 | `EmailAddressDetails.recipientType` は予定出席者には `undefined` を返します。 |
| ホストのバージョン文字列 | 実際のホストの種類によって異なる [diagnostics.hostVersion](/javascript/api/outlook/office.diagnostics#hostversion) が返すバージョン文字列の形式。例:<ul><li>Windows 上の Outlook:`15.0.4454.1002`</li><li>Web 上の Outlook:`15.0.918.2`</li></ul> |Outlook on the Mac で返される`Diagnostics.hostVersion`バージョン文字列の例を次に示します。`15.0 (140325)` |
| アイテムのカスタム プロパティ | ネットワークが使用できなくなっても、アドインはキャッシュに入っているカスタム プロパティに引き続きアクセスできます。 | Outlook on Mac はカスタムプロパティをキャッシュに入れないので、ネットワークがダウンした場合、アドインはアクセスできなくなります。 |
| 添付ファイルの詳細 | [AttachmentDetails](/javascript/api/outlook/office.attachmentdetails) オブジェクト内のコンテンツ タイプと添付ファイルの名前は、ホストの種類によって異なります。<ul><li>`AttachmentDetails.contentType` の JSON 例: `"contentType": "image/x-png"`。 </li><li>`AttachmentDetails.name` にはファイル名拡張子は含まれません。たとえば、添付ファイルが「RE: Summer activity」という件名のメッセージの場合、添付ファイル名を表す JSON オブジェクトは `"name": "RE: Summer activity"` になります。</li></ul> | <ul><li>`AttachmentDetails.contentType` の JSON 例: `"contentType" "image/png"`</li><li>`AttachmentDetails.name` には、ファイル名拡張子が必ず含まれます。メール アイテムの添付ファイルの拡張子は .eml で、予定の拡張子は .ics です。添付ファイルが「RE: Summer activity」という件名の電子メールである場合、その添付ファイル名を表す JSON オブジェクトは `"name": "RE: Summer activity.eml"` になります。<p>**注**: アドインを介するなど、ファイルがプログラムによって拡張子なしで添付される場合、`AttachmentDetails.name` にはファイル名の一部として拡張子は含まれません。</p></li></ul> |
| `dateTimeCreated` と `dateTimeModified` のプロパティでタイム ゾーンを表す文字列 |例: `Thu Mar 13 2014 14:09:11 GMT+0800 (China Standard Time)` | 例: `Thu Mar 13 2014 14:09:11 GMT+0800 (CST)` |
| `dateTimeCreated` と `dateTimeModified` の時間精度 | 次に示すコードをアドインで使用している場合、最大の精度はミリ秒単位になります:<br/>`JSON.stringify(Office.context.mailbox.item, null, 4);`| 精度は最高で秒単位となります。 |

