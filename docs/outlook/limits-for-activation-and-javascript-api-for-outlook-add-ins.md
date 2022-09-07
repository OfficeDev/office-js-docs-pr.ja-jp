---
title: Outlook アドインでのアクティブ化と API の使用の制限
description: 特定のアクティブ化と API の使用方法のガイドラインを確認し、アドインをこれらの制限内で実装します。
ms.date: 08/30/2022
ms.localizationpriority: medium
ms.openlocfilehash: b09886e49b0d980dbbf2465df7d077cd16a04f4d
ms.sourcegitcommit: 889d23061a9413deebf9092d675655f13704c727
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/07/2022
ms.locfileid: "67616015"
---
# <a name="limits-for-activation-and-javascript-api-for-outlook-add-ins"></a>Outlook アドインのアクティブ化と JavaScript API の制限

Outlook アドインのユーザーに満足のいくエクスペリエンスを提供するには、特定のアクティブ化ルールと API の使用に関するガイドラインを理解し、制限の範囲内に収まるようにアドインを実装する必要があります。 これらのガイドラインは、個々のアドインがExchange Serverまたは Outlook に対して、ライセンス認証ルールや Office JavaScript API への呼び出しを処理するために異常に長い時間を費やすことを要求できないようにするために存在し、Outlook やその他のアドインの全体的なユーザー エクスペリエンスに影響を与えます。これらの制限は、アドイン マニフェストでアクティブ化ルールを設計し、カスタム プロパティ、ローミング設定、受信者、Exchange Web Services (EWS) の要求と応答、非同期呼び出しを使用する場合に適用されます。

> [!NOTE]
> アドインが Outlook リッチ クライアントで実行されている場合は、特定のランタイム リソースの使用制限内でアドインが実行されることを確認する必要もあります。

## <a name="limits-on-where-add-ins-activate"></a>アドインのアクティブ化の制限

アドインが実行する場所とアクティブ化しない場所の詳細については、Outlook アドインの概要ページの「 [アドインで使用可能なメールボックス アイテム](outlook-add-ins-overview.md#mailbox-items-available-to-add-ins) 」セクションを参照してください。

## <a name="limits-for-activation-rules"></a>アクティブ化ルールの制限

Outlook アドインのアクティブ化ルールを設計する際には、以下のガイドラインに従います。

- マニフェストのサイズを 256 KB までに制限します。 その制限を超えた場合、Exchange メールボックスの Outlook アドインをインストールすることはできません。

- アドインで指定できるアクティブ化ルールの数は最大 15 です。 その制限を超えた場合、アドインをインストールすることはできません。

- 選択したアイテムの本文に [ItemHasKnownEntity](/javascript/api/manifest/rule#itemhasknownentity-rule) ルールを使用する場合、Outlook リッチ クライアントでは、このルールを本文の最初の 1 MB のみに適用し、この制限を超えた本文の残りの部分には適用しません。 本文の最初の MB の後にのみ一致する場合、アドインはアクティブになりません。 その可能性が高い場合は、アクティブ化の条件を再設計してください。

- または [ItemHasRegularExpressionMatch](/javascript/api/manifest/rule#itemhasregularexpressionmatch-rule) 規則で`ItemHasKnownEntity`正規表現を使用する場合は、Outlook アプリケーションに一般的に適用される次の制限とガイドライン、およびアプリケーションによって異なる表 1、2、および 3 で説明されているものに注意してください。
  - アドインのアクティブ化規則では、最大 5 つの正規表現のみを指定します。 その制限を超えた場合、アドインをインストールすることはできません。
  - 最初の 50 個の一致内のメソッド呼び出しによって `getRegExMatches` 予期される結果が返されるように正規表現を指定します。
  - **重要**: 文字列は、正規表現に一致した結果の文字列に基づいて強調表示されます。 ただし、強調表示された出現箇所は、負の先読み`(?!text)`、ルックビハインド、負のルックビハ`(?<=text)``(?<!text)`インドなどの実際の正規表現アサーションの結果と完全には一致しない可能性があります。 たとえば、"Like under,under score, and アンダースコア" の正規表現 `under(?!score)` を使用すると、最初の 2 つの文字列だけでなく、すべての出現箇所で文字列 "under" が強調表示されます。

表 1 に制限を示し、Outlook リッチ クライアントとOutlook on the webまたはモバイル デバイスの正規表現のサポートの違いについて説明します。 デバイスやアイテムの本文の種類によってサポートが異なることはありません。

**表 1.正規表現のサポートの一般的な違い**

|Outlook リッチ クライアント|Outlook on the web またはモバイル デバイス|
|:-----|:-----|
|Visual Studio の標準テンプレート ライブラリの一部として提供されている C++ 正規表現エンジンを使用します。 このエンジンは ECMAScript 5 標準に準拠しています。 |JavaScript の一部である正規表現評価を使用します。これはブラウザーによって提供されるものであり、ECMAScript 5 のスーパーセットをサポートしています。|
|正規表現エンジンが異なるため、定義済みの文字クラスに基づくカスタム文字クラスを含む正規表現を使用すると、Outlook リッチ クライアントでは、Outlook on the webまたはモバイル デバイスとは異なる結果が返されます。<br/><br/>たとえば、正規表現 `[\s\S]{0,100}` は、空白または空白以外の 1 文字の 0 ~ 100 の任意の数値と一致します。 この正規表現は、Outlook on the webおよびモバイル デバイスとは異なる結果を Outlook リッチ クライアントで返します。<br/><br/>回避策として `(\s\|\S){0,100}` 正規表現を書き直す必要があります。 このように書き換えると、空白文字または空白文字以外の任意の数 (0 から 100) と一致します。<br/><br/>各 Outlook クライアントで各正規表現を徹底的にテストする必要があります。正規表現が異なる結果を返す場合は、正規表現を書き換えてください。 |各 Outlook クライアントで各正規表現を徹底的にテストする必要があります。正規表現が異なる結果を返す場合は、正規表現を書き換えてください。|
|既定では、アドインのすべての正規表現の評価は 1 秒に制限されています。 この制限を超えると、再評価が最大 3 回試行されます。 再評価の制限を超えて、Outlook リッチ クライアントは、いずれかの Outlook クライアントで同じメールボックスに対してアドインを実行できないようにします。<br/><br/>管理者は、レジストリ キーと`OutlookActivationManagerRetryLimit`レジストリ キーを使用して、これらの評価制限を`OutlookActivationAlertThreshold`オーバーライドできます。|Outlook リッチ クライアントと同じリソース監視またはレジストリ設定をサポートしないでください。 ただし、Outlook リッチ クライアントで過剰な評価時間を必要とする正規表現を持つアドインは、すべての Outlook クライアントで同じメールボックスに対して無効になります。|

表 2 に制限事項を示します。また、Outlook のそれぞれで正規表現を適用するアイテムの本文での違いについても説明します。正規表現がアイテムの本文に適用される場合、デバイスやアイテムの本文の種類によって制限が異なることがあります。

**表 2評価対象アイテムの本文のサイズ制限**

||Outlook リッチ クライアント|モバイル デバイス上の Outlook|Outlook on the web|
|:-----|:-----|:-----|:-----|
|**フォーム ファクター**|サポートされている任意のデバイス。|Android スマートフォン、iPad、または iPhone。|Android スマートフォン、iPad、iPhone 以外でサポートされているデバイス。|
|**プレーン テキスト アイテムの本文**|RegEx は本文のデータの最初の 1 MB に適用されますが、制限を超える残りの本文には適用されません。|本文が 16,000 文字未満の場合にのみアドインがアクティブ化されます。|本文が 500,000 文字未満の場合にのみアドインがアクティブ化されます。|
|**HTML アイテムの本文**|RegEx は本文のデータの最初の 512 KB に適用されますが、制限を超える残りの本文には適用されません (実際の文字数はエンコードによって異なり、1 文字あたり 1 ～ 4 バイトの範囲でばらつくことがあります)。|RegEx は最初の 64,000 文字 (HTML タグ文字を含む) に適用されますが、制限を超える残りの本文には適用されません。|本文が 500,000 文字未満の場合にのみアドインがアクティブ化されます。|

表 3 に制限を示し、各 Outlook クライアントが正規表現を評価した後に返す一致の違いについて説明します。 サポートは特定の種類のデバイスに依存しませんが、アイテム本体に正規表現が適用されている場合は、アイテム本文の種類に依存する場合があります。

**表 3返される一致の制限**

||Outlook リッチ クライアント|Outlook on the web またはモバイル デバイス|
|:-----|:-----|:-----|
|**一致が返される順序**|Outlook リッチ クライアントで、Outlook on the webまたはモバイル デバイスで同じアイテムに適用されているのと同じ正規表現の一致が異なっているとします`getRegExMatches`。|Outlook リッチ クライアントでは、Outlook on the webデバイスやモバイル デバイスとは異なる順序で一致するとします`getRegExMatches`。|
|**プレーン テキスト アイテムの本文**|`getRegExMatches` は、最大 50 個の一致に対して最大 1,536 文字 (1.5 KB) の一致を返します。<br/><br/>**注**: `getRegExMatches` 返された配列内の特定の順序で一致を返しません。 一般に、同じアイテムに適用された同じ正規表現に対する Outlook リッチ クライアントでの一致の順序が、Outlook on the webおよびモバイル デバイスの場合とは異なるとします。|`getRegExMatches` は、最大 50 個の一致に対して最大 3,072 文字 (3 KB) の一致を返します。|
|**HTML アイテムの本文**|`getRegExMatches` は、最大 50 個の一致に対して最大 3,072 文字 (3 KB) の一致を返します。<br/> <br/> **注**: `getRegExMatches` 返される配列内の特定の順序で一致を返しません。 一般に、同じアイテムに適用された同じ正規表現に対する Outlook リッチ クライアントでの一致の順序が、Outlook on the webおよびモバイル デバイスの場合とは異なるとします。|`getRegExMatches` は、最大 50 個の一致に対して最大 3,072 文字 (3 KB) の一致を返します。|

## <a name="limits-for-javascript-api"></a>JavaScript API の制限

アクティブ化規則に関する上記のガイドラインとは別に、各 Outlook クライアントは、表 4 で説明されているように、JavaScript オブジェクト モデルに特定の制限を適用します。

**表 4.Office JavaScript API を使用して特定のデータを取得または設定するための制限**

|機能|制限|関連する API|説明|
|:-----|:-----|:-----|:-----|
|カスタム プロパティ|2,500 文字|[CustomProperties](/javascript/api/outlook/office.customproperties) オブジェクト<br/> <br/>[item.loadCustomPropertiesAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods) メソッド|予定またはメッセージのアイテムのすべてのカスタム プロパティに関する制限です。 アドインのすべてのカスタム プロパティの合計サイズがこの制限を超えた場合、すべての Outlook クライアントはエラーを返します。|
|ローミングの設定|文字数 32 KB|[RoamingSettings](/javascript/api/outlook/office.roamingsettings) オブジェクト<br/><br/> [context.roamingSettings](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context#properties) プロパティ|アドインのすべてのローミング設定に関する制限です。 設定がこの制限を超えると、すべての Outlook クライアントからエラーが返されます。|
|インターネット ヘッダー|Exchange Onlineのメッセージあたり 256 KB<br/><br/>Exchange オンプレミスの組織の管理者によって決定されるヘッダー サイズの制限|[InternetHeaders.setAsync](/javascript/api/outlook/office.internetheaders) メソッド|メッセージに適用できるヘッダーの合計サイズ制限。|
|よく知られているエンティティの抽出|2,000 文字|[item.getEntities](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods) メソッド<br/> <br/>[item.getEntitiesByType](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods) メソッド<br/> <br/>[item.getFilteredEntitiesByName](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods) メソッド|Exchange Server でアイテム本体上のよく知られたエンティティを抽出する際の制限。 Exchange Server では、その制限を超えるエンティティが無視されます。 この制限は、アドインがルールを使用 `ItemHasKnownEntity` するかどうかに依存しません。|
|Exchange Web サービス|文字数 1 MB|[mailbox.makeEwsRequestAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) メソッド|呼び出しに対する要求または応答の `Mailbox.makeEwsRequestAsync` 制限。|
|受信者|100 の受信者|[item.requiredAttendees](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) プロパティ<br/> <br/>[item.optionalAttendees](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) プロパティ<br/> <br/>[item.to](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) プロパティ<br/> <br/>[item.cc](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) プロパティ<br/> <br/>[Recipients.addAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-addasync-member(1)) メソッド<br/> <br/>[Recipient.getAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-getasync-member(1)) メソッド<br/> <br/>[Recipient.setAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-setasync-member(1)) メソッド|各プロパティで指定された受信者の制限|
|表示名|255 文字|[EmailAddressDetails.displayName](/javascript/api/outlook/office.emailaddressdetails#outlook-office-emailaddressdetails-displayname-member) プロパティ<br/><br/> [Recipients](/javascript/api/outlook/office.recipients) オブジェクト<br/><br/> `item.requiredAttendees` プロパティ<br/><br/> `item.optionalAttendees` プロパティ <br/><br/>`item.to` プロパティ <br/><br/>`item.cc` プロパティ|予定やメッセージの表示名の長さの制限。|
|件名の設定|255 文字|[Mailbox.displayNewAppointmentForm](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) メソッド<br/><br/> [Subject.setAsync](/javascript/api/outlook/office.subject#outlook-office-subject-setasync-member(1)) メソッド|新しい予定のフォームの件名、または予定やメッセージの件名の設定に関する制限。|
|場所の設定|255 文字|[Location.setAsync](/javascript/api/outlook/office.location#outlook-office-location-setasync-member(1)) メソッド|予定または会議出席依頼の場所の設定に関する制限。|
|新しい予定のフォームの本文|文字数 32 KB|`Mailbox.displayNewAppointmentForm` メソッド|新しい予定のフォームでの本文に関する制限。|
|既存のアイテムの本文の表示|文字数 32 KB|[mailbox.displayAppointmentForm](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) メソッド<br/><br/> [mailbox.displayMessageForm](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) メソッド|Outlook on the webおよびモバイル デバイスの場合: 既存の予定またはメッセージ フォームの本文の制限。|
|本文の設定|文字数 1 MB|[Body.prependAsync](/javascript/api/outlook/office.body#outlook-office-body-prependasync-member(1)) メソッド<br/> <br/>[Body.setAsync](/javascript/api/outlook/office.body#outlook-office-body-setasync-member(1))<br/><br/>[Body.setSelectedDataAsync](/javascript/api/outlook/office.body#outlook-office-body-setselecteddataasync-member(1)) メソッド|予定またはメッセージ アイテムの本文の設定に関する制限。|
|添付ファイルの数|Outlook on the webおよびモバイル デバイス上の 499 ファイル|[item.addFileAttachmentAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods) メソッド|アイテムを送信するために添付できるファイル数の制限。 Outlook on the webデバイスとモバイル デバイスでは、通常、ユーザー インターフェイスと `addFileAttachmentAsync`. Outlook リッチ クライアントでは、添付ファイルの数は特に制限されません。 ただし、すべての Outlook クライアントでは、ユーザーのExchange Serverが構成されている添付ファイルのサイズの制限が適用されます。 "添付ファイルのサイズ" については、次の行を参照してください。|
|添付ファイルのサイズ|Exchange Server に依存|`item.addFileAttachmentAsync` メソッド|管理者がユーザーのメールボックスの Exchange Server で構成できるアイテムの、すべての添付ファイルのサイズには制限があります。Outlook リッチ クライアントの場合、アイテムの添付ファイルの数が制限されます。 Outlook on the webデバイスとモバイル デバイスの場合、添付ファイルの数とすべての添付ファイルのサイズという 2 つの制限のうち小さいほど、アイテムの実際の添付ファイルが制限されます。|
|添付ファイルの名前|255 文字|`item.addFileAttachmentAsync` メソッド|アイテムに追加する添付ファイルのファイル名の長さを制限します。|
|添付ファイルの URI|2048 文字|`item.addFileAttachmentAsync` メソッド|アイテムに添付ファイルとして追加するファイル名の URI に関する制限。|
|添付ファイルの ID|100 文字|[item.addItemAttachmentAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods) メソッド<br/><br/> [item.removeAttachmentAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods) メソッド|アイテムに追加またはアイテムから削除する添付ファイルの ID の長さに関する制限。|
|非同期呼び出し|3 つの呼び出し|`item.addFileAttachmentAsync` メソッド<br/><br/>`item.addItemAttachmentAsync` メソッド<br/><br/><br/>`item.removeAttachmentAsync` メソッド<br/><br/> [Body.getTypeAsync](/javascript/api/outlook/office.body#outlook-office-body-gettypeasync-member(1)) メソッド<br/><br/>`Body.prependAsync` メソッド<br/><br/>`Body.setSelectedDataAsync` メソッド<br/><br/> [CustomProperties.saveAsync](/javascript/api/outlook/office.customproperties#outlook-office-customproperties-saveasync-member(1)) メソッド<br/><br/><br/> [項目。LoadCustomPropertiesAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods) メソッド<br/><br/><br/> [Location.getAsync](/javascript/api/outlook/office.location#outlook-office-location-getasync-member(1)) メソッド<br/><br/>`Location.setAsync` メソッド<br/><br/> [mailbox.getCallbackTokenAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) メソッド<br/><br/> [mailbox.getUserIdentityTokenAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) メソッド<br/><br/> [mailbox.makeEwsRequestAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) メソッド<br/><br/>`Recipients.addAsync` メソッド<br/><br/> [Recipients.getAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-getasync-member(1)) メソッド<br/><br/>`Recipients.setAsync` メソッド<br/><br/> [RoamingSettings.saveAsync](/javascript/api/outlook/office.roamingsettings#outlook-office-roamingsettings-saveasync-member(1)) メソッド<br/><br/> [Subject.getAsync](/javascript/api/outlook/office.subject#outlook-office-subject-getasync-member(1)) メソッド<br/><br/>`Subject.setAsync` メソッド<br/><br/> [Time.getAsync](/javascript/api/outlook/office.time#outlook-office-time-getasync-member(1)) メソッド<br/><br/> [Time.setAsync](/javascript/api/outlook/office.time#outlook-office-time-setasync-member(1)) メソッド|Outlook on the webまたはモバイル デバイスの場合: ブラウザーではサーバーへの非同期呼び出しの数が制限されているため、同時に非同期呼び出しの数を制限します。 |

## <a name="see-also"></a>関連項目

- [テスト用に Outlook アドインを展開してインストールする](testing-and-tips.md)
- [Outlook アドインに関するプライバシー、アクセス許可、セキュリティ](../concepts/privacy-and-security.md)
