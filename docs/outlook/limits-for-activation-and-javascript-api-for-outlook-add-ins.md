---
title: Outlook アドインでのアクティブ化と API の使用の制限
description: 特定のアクティブ化と API の使用方法のガイドラインを確認し、アドインをこれらの制限内で実装します。
ms.date: 06/11/2021
localization_priority: Normal
ms.openlocfilehash: eacdc0232202fd74fdd46a835bed6af5a760e7b1
ms.sourcegitcommit: 0bf0e076f705af29193abe3dba98cbfcce17b24f
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/18/2021
ms.locfileid: "53007777"
---
# <a name="limits-for-activation-and-javascript-api-for-outlook-add-ins"></a>Outlook アドインのアクティブ化と JavaScript API の制限

Outlook アドインのユーザーに満足のいくエクスペリエンスを提供するには、特定のアクティブ化ルールと API の使用に関するガイドラインを理解し、制限の範囲内に収まるようにアドインを実装する必要があります。 これらのガイドラインは、個々のアドインが Exchange Server または Outlook が Office JavaScript API のアクティブ化ルールまたは呼び出しを処理するために異常に長い時間を費やす必要が生じないので、Outlook や他のアドインの全体的なユーザー エクスペリエンスに影響を及ぼします。これらの制限は、アドイン マニフェストでアクティブ化ルールを設計し、カスタム プロパティ、ローミング設定、受信者、Exchange Web サービス (EWS) 要求と応答、および非同期呼び出しを使用する場合に適用されます。

> [!NOTE]
> アドインを Outlook リッチ クライアントで実行する場合は、そのアドインが一定のランタイム リソース使用制限の範囲内で実行されているかを確認する必要もあります。

## <a name="limits-on-where-add-ins-activate"></a>アドインのアクティブ化の制限

アドインがアクティブ化する場所とアクティブ化しない場所の詳細については、「Outlook アドイン[](outlook-add-ins-overview.md#mailbox-items-available-to-add-ins)の概要」ページの「アドインで使用可能なメールボックス アイテム」セクションを参照してください。

## <a name="limits-for-activation-rules"></a>アクティブ化ルールの制限

Outlook アドインのアクティブ化ルールを設計する際には、以下のガイドラインに従います。

- マニフェストのサイズを 256 KB までに制限します。この上限を超える場合は、Exchange メールボックス用にその Outlook アドインをインストールすることはできません。

- アドインで指定できるアクティブ化ルールの数は最大 15 です。この上限を超える場合、そのアドインはインストールできません。

- 選択したアイテムの本文に [ItemHasKnownEntity](../reference/manifest/rule.md#itemhasknownentity-rule) ルールを使用する場合、Outlook リッチ クライアントでは、このルールを本文の最初の 1 MB のみに適用し、この制限を超えた本文の残りの部分には適用しません。本文の最初の 1 MB の後にしか一致するものが存在しない場合、アドインはアクティブにはなりません。その可能性が高い場合は、アクティブ化の条件を再設計してください。

- `ItemHasKnownEntity`[または ItemHasRegularExpressionMatch](../reference/manifest/rule.md#itemhasregularexpressionmatch-rule)ルールで正規表現を使用する場合は、Outlook アプリケーションに一般的に適用される次の制限とガイドライン、およびアプリケーションによって異なる表 1、2、3 に記載されている制限とガイドラインに注意してください。
  - アドインで指定できるアクティブ化ルールの正規表現は最大 5 つです。その制限を超えた場合は、アドインをインストールできません。
  - 予想される結果が、最初の 50 回の一致内のメソッド呼び出しによって返される正規表現 `getRegExMatches` を指定します。
  - 正規表現で先読みアサーションは指定できますが、後読み `(?<=text)` および否定の後読み `(?<!text)` アサーションは指定できません。

表 1 に、制限を示し、リッチ クライアントとモバイル デバイスの間の正規表現のサポートOutlookを示Outlook on the web示します。 デバイスやアイテムの本文の種類によってサポートが異なることはありません。

**表 1.正規表現のサポートの一般的な違い**

|Outlook リッチ クライアント|Outlook on the web またはモバイル デバイス|
|:-----|:-----|
|Visual Studio の標準テンプレート ライブラリの一部として提供されている C++ 正規表現エンジンを使用します。 このエンジンは ECMAScript 5 標準に準拠しています。 |JavaScript の一部である正規表現評価を使用します。これはブラウザーによって提供されるものであり、ECMAScript 5 のスーパーセットをサポートしています。|
|正規表現エンジンが異なっているので、定義済みの文字クラスに基づくカスタム文字クラスを含む正規表現は、Outlook on the web またはモバイル デバイスと異なる Outlook リッチ クライアントで異なる結果を返す可能性があります。<br/><br/>たとえば、正規表現 `[\s\S]{0,100}` は、空白文字または空白以外の単一文字の任意の数 (0 から 100) と一致します。 この正規表現は、リッチ クライアントとモバイル デバイスOutlook異なる結果Outlook on the web返します。<br/><br/>回避策としては、正規表現を `(\s\|\S){0,100}` として書き換えてください。 このように書き換えると、空白文字または空白文字以外の任意の数 (0 から 100) と一致します。<br/><br/>各正規表現は、クライアントの各Outlookテストし、正規表現が異なる結果を返す場合は、正規表現を書き換える必要があります。 |各正規表現は、クライアントの各Outlookテストし、正規表現が異なる結果を返す場合は、正規表現を書き換える必要があります。|
|既定では、アドインのすべての正規表現の評価は 1 秒に制限されています。 この制限を超えると、再評価が最大 3 回試行されます。 再評価の制限を超えて、Outlookリッチ クライアントは、任意のクライアントで同じメールボックスでアドインを実行Outlookします。<br/><br/>管理者は、レジストリ キーとレジストリ キーを使用して、これらの評価 `OutlookActivationAlertThreshold` 制限 `OutlookActivationManagerRetryLimit` を上書きできます。|Outlook リッチ クライアントと同じリソース監視設定およびレジストリ設定はサポートしていません。 ただし、正規表現を使用するアドインでは、Outlook リッチ クライアントで過剰な評価時間が必要な場合、すべてのクライアントで同じメールボックスOutlookされます。|

表 2 に制限事項を示します。また、Outlook のそれぞれで正規表現を適用するアイテムの本文での違いについても説明します。正規表現がアイテムの本文に適用される場合、デバイスやアイテムの本文の種類によって制限が異なることがあります。

**表 2評価対象アイテムの本文のサイズ制限**

||Outlook リッチ クライアント|Outlookデバイス上での設定|Outlook on the web|
|:-----|:-----|:-----|:-----|
|**フォーム ファクター**|サポートされる任意のデバイス。|Android スマートフォン、iPad、または iPhone|Android スマートフォン、iPad、および iPhone 以外のサポートされている任意のデバイス|
|**プレーン テキスト アイテムの本文**|RegEx は本文のデータの最初の 1 MB に適用されますが、制限を超える残りの本文には適用されません。|本文が 16,000 文字未満の場合にのみアドインがアクティブ化されます。|本文が 500,000 文字未満の場合にのみアドインがアクティブ化されます。|
|**HTML アイテムの本文**|RegEx は本文のデータの最初の 512 KB に適用されますが、制限を超える残りの本文には適用されません (実際の文字数はエンコードによって異なり、1 文字あたり 1 ～ 4 バイトの範囲でばらつくことがあります)。|RegEx は最初の 64,000 文字 (HTML タグ文字を含む) に適用されますが、制限を超える残りの本文には適用されません。|本文が 500,000 文字未満の場合にのみアドインがアクティブ化されます。|

表 3 に、制限の一覧を示し、正規表現を評価した後Outlookクライアントが返す一致の違いを示します。 サポートは、デバイスの特定の種類に依存しますが、アイテムの本文に正規表現が適用されている場合は、アイテムの本文の種類に依存する場合があります。

**表 3返される一致の制限**

||Outlook リッチ クライアント|Outlook on the web またはモバイル デバイス|
|:-----|:-----|:-----|
|**一致が返される順序**|同じアイテムに適用される同じ正規表現の一致が、Outlookまたはモバイル デバイスの場合と異なOutlook on the web `getRegExMatches` 想定します。|リッチ クライアント内の一致する順序は、Outlookモバイル デバイスと異なる順序 `getRegExMatches` Outlook on the web想定します。|
|**プレーン テキスト アイテムの本文**|`getRegExMatches` 最大 50 一致の場合、最大 1,536 (1.5 KB) 文字の一致を返します。<br/><br/>**注**: `getRegExMatches` 返される配列内の特定の順序で一致する値は返されません。 一般に、同じアイテムに適用される同じ正規表現に対する Outlook リッチ クライアントでの一致の順序が、Outlook on the web デバイスとモバイル デバイスで異なると仮定します。|`getRegExMatches` 最大 3,072 (3 KB) 文字の一致を返し、最大 50 件の一致を返します。|
|**HTML アイテムの本文**|`getRegExMatches` 最大 3,072 (3 KB) 文字の一致を返し、最大 50 件の一致を返します。<br/> <br/> **注**: `getRegExMatches` 返される配列内の特定の順序で一致する値は返されません。 一般に、同じアイテムに適用される同じ正規表現に対する Outlook リッチ クライアントでの一致の順序が、Outlook on the web デバイスとモバイル デバイスで異なると仮定します。|`getRegExMatches` 最大 3,072 (3 KB) 文字の一致を返し、最大 50 件の一致を返します。|

## <a name="limits-for-javascript-api"></a>JavaScript API の制限

ライセンス認証ルールに関する前述のガイドラインを除き、各 Outlook クライアントは、表 4 で説明したように、JavaScript オブジェクト モデルに特定の制限を適用します。

**表 4.JavaScript API を使用して特定のデータを取得または設定Office制限**

|機能|制限|関連する API|説明|
|:-----|:-----|:-----|:-----|
|カスタム プロパティ|2,500 文字|[CustomProperties](/javascript/api/outlook/office.CustomProperties) オブジェクト<br/> <br/>[item.loadCustomPropertiesAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) メソッド|予定またはメッセージのアイテムのすべてのカスタム プロパティに関する制限です。 アドインのすべてのOutlookプロパティの合計サイズがこの制限を超えると、すべてのクライアントがエラーを返します。|
|ローミングの設定|文字数 32 KB|[RoamingSettings](/javascript/api/outlook/office.RoamingSettings) オブジェクト<br/><br/> [context.roamingSettings](../reference/objectmodel/preview-requirement-set/office.context.md#properties) プロパティ|アドインのすべてのローミング設定に関する制限です。 設定がOutlookを超えると、すべてのクライアントがエラーを返します。|
|よく知られているエンティティの抽出|2,000 文字|[item.getEntities](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) メソッド<br/> <br/>[item.getEntitiesByType](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) メソッド<br/> <br/>[item.getFilteredEntitiesByName](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) メソッド|Exchange Server でアイテム本体上のよく知られたエンティティを抽出する際の制限。 Exchange Server では、その制限を超えるエンティティが無視されます。 この制限は、アドインがルールを使用するかどうかに依存しない点に注意 `ItemHasKnownEntity` してください。|
|Exchange Web サービス|文字数 1 MB|[mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) メソッド|呼び出しに対する要求または応答の `Mailbox.makeEwsRequestAsync` 制限。|
|受信者|100 の受信者|[item.requiredAttendees](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) プロパティ<br/> <br/>[item.optionalAttendees](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) プロパティ<br/> <br/>[item.to](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) プロパティ<br/> <br/>[item.cc](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) プロパティ<br/> <br/>[Recipients.addAsync](/javascript/api/outlook/office.Recipients#addasync-recipients--options--callback-) メソッド<br/> <br/>[Recipient.getAsync](/javascript/api/outlook/office.Recipients#getasync-options--callback-) メソッド<br/> <br/>[Recipient.setAsync](/javascript/api/outlook/office.Recipients#setasync-recipients--options--callback-) メソッド|各プロパティで指定された受信者の制限|
|表示名|255 文字|[EmailAddressDetails.displayName](/javascript/api/outlook/office.emailaddressdetails#displayname) プロパティ<br/><br/> [Recipients](/javascript/api/outlook/office.Recipients) オブジェクト<br/><br/> `item.requiredAttendees` プロパティ<br/><br/> `item.optionalAttendees` プロパティ <br/><br/>`item.to` プロパティ <br/><br/>`item.cc` プロパティ|予定やメッセージの表示名の長さの制限。|
|件名の設定|255 文字|[Mailbox.displayNewAppointmentForm](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) メソッド<br/><br/> [Subject.setAsync](/javascript/api/outlook/office.Subject#setasync-subject--options--callback-) メソッド|新しい予定のフォームの件名、または予定やメッセージの件名の設定に関する制限。|
|場所の設定|255 文字|[Location.setAsync](/javascript/api/outlook/office.Location#setasync-location--options--callback-) メソッド|予定または会議出席依頼の場所の設定に関する制限。|
|新しい予定のフォームの本文|文字数 32 KB|`Mailbox.displayNewAppointmentForm` メソッド|新しい予定のフォームでの本文に関する制限。|
|既存のアイテムの本文の表示|文字数 32 KB|[mailbox.displayAppointmentForm](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) メソッド<br/><br/> [mailbox.displayMessageForm](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) メソッド|モバイル Outlook on the webデバイスの場合: 既存の予定またはメッセージ フォーム内の本文の制限。|
|本文の設定|文字数 1 MB|[Body.prependAsync](/javascript/api/outlook/office.Body#prependasync-data--options--callback-) メソッド<br/> <br/>[Body.setAsync](/javascript/api/outlook/office.Body#setasync-data--options--callback-)<br/><br/>[Body.setSelectedDataAsync](/javascript/api/outlook/office.Body#setselecteddataasync-data--options--callback-) メソッド|予定またはメッセージ アイテムの本文の設定に関する制限。|
|添付ファイルの数|モバイル デバイスとモバイル Outlook on the web 499 ファイル|[item.addFileAttachmentAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) メソッド|アイテムを送信するために添付できるファイル数の制限。 Outlook on the webモバイル デバイスでは、通常、ユーザー インターフェイスを通じて最大 499 のファイルの添付を制限します `addFileAttachmentAsync` 。 Outlook リッチ クライアントでは、添付ファイルの数は特に制限されません。 ただし、すべてのOutlookクライアントは、ユーザーが構成した添付ファイルのサイズExchange Serverを監視します。 "添付ファイルのサイズ" については、次の行を参照してください。|
|添付ファイルのサイズ|Exchange Server に依存|`item.addFileAttachmentAsync` メソッド|管理者がユーザーのメールボックスの Exchange Server で構成できるアイテムの、すべての添付ファイルのサイズには制限があります。Outlook リッチ クライアントの場合、アイテムの添付ファイルの数が制限されます。 モバイル Outlook on the webでは、添付ファイルの数とすべての添付ファイルのサイズという 2 つの制限の小さい方が、アイテムの実際の添付ファイルを制限します。|
|添付ファイルの名前|255 文字|`item.addFileAttachmentAsync` メソッド|アイテムに追加する添付ファイルのファイル名の長さを制限します。|
|添付ファイルの URI|2048 文字|`item.addFileAttachmentAsync` メソッド|アイテムに添付ファイルとして追加するファイル名の URI に関する制限。|
|添付ファイルの ID|100 文字|[item.addItemAttachmentAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) メソッド<br/><br/> [item.removeAttachmentAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) メソッド|アイテムに追加またはアイテムから削除する添付ファイルの ID の長さに関する制限。|
|非同期呼び出し|3 つの呼び出し|`item.addFileAttachmentAsync` メソッド<br/><br/>`item.addItemAttachmentAsync` メソッド<br/><br/><br/>`item.removeAttachmentAsync` メソッド<br/><br/> [Body.getTypeAsync](/javascript/api/outlook/office.Body#gettypeasync-options--callback-) メソッド<br/><br/>`Body.prependAsync` メソッド<br/><br/>`Body.setSelectedDataAsync` メソッド<br/><br/> [CustomProperties.saveAsync](/javascript/api/outlook/office.CustomProperties#saveasync-callback--asynccontext-) メソッド<br/><br/><br/> [item.LoadCustomPropertiesAysnc](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) メソッド<br/><br/><br/> [Location.getAsync](/javascript/api/outlook/office.Location#getasync-options--callback-) メソッド<br/><br/>`Location.setAsync` メソッド<br/><br/> [mailbox.getCallbackTokenAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) メソッド<br/><br/> [mailbox.getUserIdentityTokenAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) メソッド<br/><br/> [mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) メソッド<br/><br/>`Recipients.addAsync` メソッド<br/><br/> [Recipients.getAsync](/javascript/api/outlook/office.Recipients#getasync-options--callback-) メソッド<br/><br/>`Recipients.setAsync` メソッド<br/><br/> [RoamingSettings.saveAsync](/javascript/api/outlook/office.RoamingSettings#saveasync-callback-) メソッド<br/><br/> [Subject.getAsync](/javascript/api/outlook/office.Subject#getasync-options--callback-) メソッド<br/><br/>`Subject.setAsync` メソッド<br/><br/> [Time.getAsync](/javascript/api/outlook/office.Time#getasync-options--callback-) メソッド<br/><br/> [Time.setAsync](/javascript/api/outlook/office.Time#setasync-datetime--options--callback-) メソッド|モバイル Outlook on the webの場合:ブラウザーはサーバーに対する非同期呼び出しの数が限られているので、同時非同期呼び出しの数を一度に制限します。 |

## <a name="see-also"></a>関連項目

- [テスト用に Outlook アドインを展開してインストールする](testing-and-tips.md)
- [Outlook アドインに関するプライバシー、アクセス許可、セキュリティ](../concepts/privacy-and-security.md)
