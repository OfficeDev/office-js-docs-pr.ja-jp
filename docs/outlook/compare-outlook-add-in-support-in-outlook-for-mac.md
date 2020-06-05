---
title: Outlook on Mac での Outlook アドインサポートの比較
description: Outlook on Mac でのアドインのサポートと他の Outlook クライアントとの比較について説明します。
ms.date: 06/04/2020
localization_priority: Normal
ms.openlocfilehash: 13022f154c05a8275f5124ce4c5310e13af5e525
ms.sourcegitcommit: 6754aa2835e57c3a95b0c513095ba4b29744f9eb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/05/2020
ms.locfileid: "44567842"
---
# <a name="compare-outlook-add-in-support-in-outlook-on-mac-with-other-outlook-clients"></a>Outlook on Mac での outlook アドインのサポートを他の Outlook クライアントと比較する

Outlook on the web、Windows、iOS、Android などの他のクライアントと同じ方法で Outlook アドインを作成して実行すると、各クライアントの JavaScript をカスタマイズする必要はありません。 通常、アドインから Office JavaScript API への同じ呼び出しは、次の表に示す領域を除き、同じように動作します。

詳細については、「[Outlook 2013 プレビューでのテスト用メール アプリの展開とインストール](testing-and-tips.md)」を参照してください。

Mac での新しい UI のサポートの詳細については、「 [New Outlook On mac](#new-outlook-on-mac-preview)」を参照してください。

| 分野 | Web 上の Outlook、Windows、およびモバイルデバイス | Outlook on Mac |
|:-----|:-----|:-----|
| サポート対象バージョンの office.js および Office アドインのマニフェスト スキーマ | Office.js および スキーマ v1.1 のすべての API。 | Office.js および スキーマ v1.1 のすべての API。<br><br>**注**: Outlook on Mac では、16.35.308 以降のビルドのみが会議の保存をサポートしています。 それ以外の場合は、 `saveAsync` 作成モードで会議から呼び出されたときにメソッドが失敗します。 回避策については、「[Office JS API を使用して Outlook for Mac で会議を下書きとして保存できない](https://support.microsoft.com/help/4505745)」を参照してください。 |
| 定期的な予定系列のインスタンス | <ul><li>定期的な系列のマスター予定または予定インスタンスのアイテム ID および他のプロパティを取得できます。</li><li>[mailbox.displayAppointmentForm](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) を使用して、定期的な系列のインスタンスまたはマスターを表示できます。</li></ul> | <ul><li>マスター予定のアイテム ID と他のプロパティを取得できますが、定期的な系列のインスタンスのアイテム ID とプロパティは取得できません。</li><li>定期的な系列のマスター予定を表示できます。アイテム ID がない場合、定期的な系列のインスタンスは表示できません。</li></ul> |
| 予定出席者の受信者の種類 | [EmailAddressDetails.recipientType](/javascript/api/outlook/office.emailaddressdetails#recipienttype) を使用して、出席者の受信者の種類を特定できます。 | `EmailAddressDetails.recipientType` は予定出席者には `undefined` を返します。 |
| ホストクライアントのバージョン文字列 | [HostVersion](/javascript/api/outlook/office.diagnostics#hostversion)によって返されるバージョン文字列の形式は、クライアントの実際の種類によって異なります。 例:<ul><li>Windows 上の Outlook:`15.0.4454.1002`</li><li>Web 上の Outlook:`15.0.918.2`</li></ul> |Outlook on the Mac で返されるバージョン文字列の例を `Diagnostics.hostVersion` 次に示します。`15.0 (140325)` |
| アイテムのカスタム プロパティ | ネットワークが使用できなくなっても、アドインはキャッシュに入っているカスタム プロパティに引き続きアクセスできます。 | Outlook on Mac はカスタムプロパティをキャッシュに入れないので、ネットワークがダウンした場合、アドインはアクセスできなくなります。 |
| 添付ファイルの詳細 | [Attachmentdetails](/javascript/api/outlook/office.attachmentdetails)オブジェクト内のコンテンツタイプと添付ファイルの名前は、クライアントの種類によって異なります。<ul><li>`AttachmentDetails.contentType` の JSON 例: `"contentType": "image/x-png"`。 </li><li>`AttachmentDetails.name` にはファイル名拡張子は含まれません。たとえば、添付ファイルが「RE: Summer activity」という件名のメッセージの場合、添付ファイル名を表す JSON オブジェクトは `"name": "RE: Summer activity"` になります。</li></ul> | <ul><li>`AttachmentDetails.contentType` の JSON 例: `"contentType" "image/png"`</li><li>`AttachmentDetails.name` には、ファイル名拡張子が必ず含まれます。メール アイテムの添付ファイルの拡張子は .eml で、予定の拡張子は .ics です。添付ファイルが「RE: Summer activity」という件名の電子メールである場合、その添付ファイル名を表す JSON オブジェクトは `"name": "RE: Summer activity.eml"` になります。<p>**注**: アドインを介するなど、ファイルがプログラムによって拡張子なしで添付される場合、`AttachmentDetails.name` にはファイル名の一部として拡張子は含まれません。</p></li></ul> |
| `dateTimeCreated` と `dateTimeModified` のプロパティでタイム ゾーンを表す文字列 |例: `Thu Mar 13 2014 14:09:11 GMT+0800 (China Standard Time)` | 例: `Thu Mar 13 2014 14:09:11 GMT+0800 (CST)` |
| `dateTimeCreated` と `dateTimeModified` の時間精度 | 次に示すコードをアドインで使用している場合、最大の精度はミリ秒単位になります:<br/>`JSON.stringify(Office.context.mailbox.item, null, 4);`| 精度は最高で秒単位となります。 |

## <a name="new-outlook-on-mac-preview"></a>新しい Outlook on Mac (プレビュー)

これで、Outlook アドインは新しい Mac UI でサポートされるようになりました。要件セットは1.6 です。 ただし、次の要件セットと機能はまだサポートされて**いません**。

1. API 要件は1.7 と1.8 を設定します。
1. Pinnable 作業ウィンドウ、 `ItemChanged` イベント
1. コンテキスト アドイン
1. 送信時
1. 共有フォルダーのサポート
1. `saveAsync`会議を作成するとき
1. シングル サインオン (SSO)

新しい Outlook on the Mac をプレビューすることをお勧めします。これは、バージョン16.38.506 から入手できます。 試す方法の詳細については、「 [Insider Fast ビルドの Outlook For Mac リリースノート](https://support.microsoft.com/office/d6347358-5613-433e-a49e-a9a0e8e0462a)」を参照してください。

どの UI バージョンを使用しているかは、次のように判断できます。

**現在の UI**

&nbsp;&nbsp;&nbsp;&nbsp;![Mac での現在の UI](../images/outlook-on-mac-classic.png)

**新しい UI (プレビュー)**

&nbsp;&nbsp;&nbsp;&nbsp;![Mac でのプレビューの新しい UI](../images/outlook-on-mac-new.png)