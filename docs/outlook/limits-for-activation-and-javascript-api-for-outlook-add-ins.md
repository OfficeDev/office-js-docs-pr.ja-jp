---
title: Outlook アドインでのアクティブ化と API の使用の制限
description: 特定のアクティブ化と API の使用方法のガイドラインを確認し、アドインをこれらの制限内で実装します。
ms.date: 05/08/2020
localization_priority: Normal
ms.openlocfilehash: ad7e572152e9dd9aeb3b07f1c545184343c4e5e8
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293878"
---
# <a name="limits-for-activation-and-javascript-api-for-outlook-add-ins"></a>Outlook アドインのアクティブ化と JavaScript API の制限

Outlook アドインのユーザーに満足のいくエクスペリエンスを提供するには、特定のアクティブ化ルールと API の使用に関するガイドラインを理解し、制限の範囲内に収まるようにアドインを実装する必要があります。 これらのガイドラインは、個々のアドインが、Exchange Server または Outlook で、アクティブ化ルールを処理したり、Office JavaScript API を呼び出したりする必要があり、Outlook やその他のアドインの全体的なユーザーの操作に影響を与えることがないようにするために存在します。これらの制限は、アドインマニフェストでのアクティブ化ルールの設計、およびカスタムプロパティ、ローミング設定、受信者、Exchange Web サービス (EWS) の要求と応答、および非同期呼び出しの使用に適用されます。

> [!NOTE]
> アドインを Outlook リッチ クライアントで実行する場合は、そのアドインが一定のランタイム リソース使用制限の範囲内で実行されているかを確認する必要もあります。

## <a name="limits-on-where-add-ins-activate"></a>アドインのアクティブ化の制限

既定では、アドインはユーザーのメインメールボックスでのみアクティブになるように設計されています。 そのため、通常、アドインは共有メールボックスではアクティブ化されず、代理アクセス、アーカイブメールボックス、またはパブリックフォルダーを使用して開かれた他のユーザーのメールボックスからのフォルダーはアクティブになりません。 ただし、 [代理人アクセスまたは共有フォルダー](delegate-access.md) をサポートするアドインはアクティブ化する必要があります。

## <a name="limits-for-activation-rules"></a>アクティブ化ルールの制限

Outlook アドインのアクティブ化ルールを設計する際には、以下のガイドラインに従います。

- マニフェストのサイズを 256 KB までに制限します。この上限を超える場合は、Exchange メールボックス用にその Outlook アドインをインストールすることはできません。

- アドインで指定できるアクティブ化ルールの数は最大 15 です。この上限を超える場合、そのアドインはインストールできません。

- 選択したアイテムの本文に [ItemHasKnownEntity](../reference/manifest/rule.md#itemhasknownentity-rule) ルールを使用する場合、Outlook リッチ クライアントでは、このルールを本文の最初の 1 MB のみに適用し、この制限を超えた本文の残りの部分には適用しません。本文の最初の 1 MB の後にしか一致するものが存在しない場合、アドインはアクティブにはなりません。その可能性が高い場合は、アクティブ化の条件を再設計してください。

- または ItemHasRegularExpressionMatch ルールで正規表現を使用する場合 `ItemHasKnownEntity` は、通常、Outlook アプリケーションに適用される以下の制限とガイドライン、およびアプリケーションに応じて異なる表1、2、3で説明されているものに注意してください。 [ItemHasRegularExpressionMatch](../reference/manifest/rule.md#itemhasregularexpressionmatch-rule)
   - アドインのアクティブ化ルールに最大5つの正規表現を指定します。 You cannot install a add-in if you exceed that limit.
   - 正規表現を指定します。これは、予期した結果が、 `getRegExMatches` 最初の50一致内のメソッド呼び出しによって返されることを示します。
   - 正規表現で先読みアサーションは指定できますが、後読み `(?<=text)` および否定の後読み `(?<!text)` アサーションは指定できません。

表1に、Outlook リッチクライアントと web またはモバイルデバイスとの間での正規表現のサポートの違いを示します。 デバイスやアイテムの本文の種類によってサポートが異なることはありません。

**表 1.正規表現のサポートの一般的な違い**

|Outlook リッチ クライアント|Outlook on the web またはモバイル デバイス|
|:-----|:-----|
|Visual Studio の標準テンプレート ライブラリの一部として提供されている C++ 正規表現エンジンを使用します。 このエンジンは ECMAScript 5 標準に準拠しています。 |JavaScript の一部である正規表現評価を使用します。これはブラウザーによって提供されるものであり、ECMAScript 5 のスーパーセットをサポートしています。|
|Regex エンジンが異なるため、定義済みの文字クラスに基づいたカスタムの文字クラスを含む regex では、outlook リッチクライアントで、web またはモバイルデバイスとは異なる結果が返されることがあります。<br/><br/>たとえば、正規表現 `[\s\S]{0,100}` は、空白文字または空白以外の単一文字の任意の数 (0 から 100) と一致します。 この regex は、outlook リッチクライアントでは、web およびモバイルデバイスとは異なる結果を返します。<br/><br/>回避策としては、正規表現を `(\s\|\S){0,100}` として書き換えてください。 このように書き換えると、空白文字または空白文字以外の任意の数 (0 から 100) と一致します。<br/><br/>各 Outlook クライアント上で各 regex を徹底的にテストし、regex が異なる結果を返す場合は regex を書き直します。 |各 Outlook クライアント上で各 regex を徹底的にテストし、regex が異なる結果を返す場合は regex を書き直します。|
|既定では、アドインのすべての正規表現の評価は 1 秒に制限されています。 この制限を超えると、再評価が最大 3 回試行されます。 Outlook リッチクライアントでは、再評価の制限を超えて、Outlook クライアントの同じメールボックスに対してアドインが実行されないようにすることができます。<br/><br/>管理者は、およびレジストリキーを使用してこれらの評価の制限を無効にすることができ `OutlookActivationAlertThreshold` `OutlookActivationManagerRetryLimit` ます。|Outlook リッチ クライアントと同じリソース監視設定およびレジストリ設定はサポートしていません。 しかし、Outlook リッチクライアントでは、多くの評価時間を必要とする正規表現を使用するアドインは、すべての Outlook クライアントで同じメールボックスに対して無効になっています。|

表 2 に制限事項を示します。また、Outlook のそれぞれで正規表現を適用するアイテムの本文での違いについても説明します。正規表現がアイテムの本文に適用される場合、デバイスやアイテムの本文の種類によって制限が異なることがあります。

**表 2評価対象アイテムの本文のサイズ制限**

||Outlook リッチ クライアント|モバイルデバイスの Outlook|Outlook on the web|
|:-----|:-----|:-----|:-----|
|**フォーム ファクター**|サポートされる任意のデバイス。|Android スマートフォン、iPad、または iPhone|Android スマートフォン、iPad、および iPhone 以外のサポートされている任意のデバイス|
|**プレーン テキスト アイテムの本文**|RegEx は本文のデータの最初の 1 MB に適用されますが、制限を超える残りの本文には適用されません。|本文が 16,000 文字未満の場合にのみアドインがアクティブ化されます。|本文が 500,000 文字未満の場合にのみアドインがアクティブ化されます。|
|**HTML アイテムの本文**|RegEx は本文のデータの最初の 512 KB に適用されますが、制限を超える残りの本文には適用されません (実際の文字数はエンコードによって異なり、1 文字あたり 1 ～ 4 バイトの範囲でばらつくことがあります)。|RegEx は最初の 64,000 文字 (HTML タグ文字を含む) に適用されますが、制限を超える残りの本文には適用されません。|本文が 500,000 文字未満の場合にのみアドインがアクティブ化されます。|

表3は、制限の一覧と、正規表現を評価した後に各 Outlook クライアントが返す一致の違いについて説明しています。 サポートは特定の種類のデバイスから独立していますが、アイテムの本文に正規表現が適用されている場合は、アイテムの本文の種類によって異なります。

**表 3返される一致の制限**

||Outlook リッチ クライアント|Outlook on the web またはモバイル デバイス|
|:-----|:-----|:-----|
|**一致が返される順序**|`getRegExMatches`Outlook リッチクライアントでは、同じアイテムに適用されているのと同じ正規表現を、web またはモバイルデバイスとは異なるものとして返します。|`getRegExMatches`Outlook リッチクライアントでは、web またはモバイルデバイスとは異なる順序で一致を返します。|
|**プレーン テキスト アイテムの本文**|`getRegExMatches` 最大 1536 (1.5 KB) 文字で、最大で50の一致が返されます。<br/><br/>**注**: `getRegExMatches` は、返される配列内の特定の順序で一致するものを返しません。 通常、Outlook リッチクライアントで同じアイテムに適用された正規表現と一致する順序は、web およびモバイルデバイスの Outlook の場合とは異なります。|`getRegExMatches` 最大 3072 (3 KB) 文字の一致するものを最大50一致として返します。|
|**HTML アイテムの本文**|`getRegExMatches` 最大 3072 (3 KB) 文字の一致するものを最大50一致として返します。<br/> <br/> **注**: `getRegExMatches` は、返される配列内の特定の順序で一致するものを返しません。 通常、Outlook リッチクライアントで同じアイテムに適用された正規表現と一致する順序は、web およびモバイルデバイスの Outlook の場合とは異なります。|`getRegExMatches` 最大 3072 (3 KB) 文字の一致するものを最大50一致として返します。|

## <a name="limits-for-javascript-api"></a>JavaScript API の制限

前述のアクティブ化ルールのガイドラインとは別に、各 Outlook クライアントでは、表4に示すように、JavaScript オブジェクトモデルに一定の制限が適用されます。

**表4Office JavaScript API を使用して特定のデータを取得または設定するための制限**

|機能|制限|関連する API|説明|
|:-----|:-----|:-----|:-----|
|カスタム プロパティ|2,500 文字|[CustomProperties](/javascript/api/outlook/office.CustomProperties) オブジェクト<br/> <br/>[item.loadCustomPropertiesAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) メソッド|予定またはメッセージのアイテムのすべてのカスタム プロパティに関する制限です。 すべての Outlook クライアントは、アドインのすべてのカスタムプロパティの合計サイズがこの制限を超えた場合にエラーを返します。|
|ローミングの設定|文字数 32 KB|[RoamingSettings](/javascript/api/outlook/office.RoamingSettings) オブジェクト<br/><br/> [context.roamingSettings](../reference/objectmodel/preview-requirement-set/office.context.md#properties) プロパティ|アドインのすべてのローミング設定に関する制限です。 設定がこの制限を超えている場合は、すべての Outlook クライアントがエラーを返します。|
|よく知られているエンティティの抽出|2,000 文字|[item.getEntities](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) メソッド<br/> <br/>[item.getEntitiesByType](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) メソッド<br/> <br/>[item.getFilteredEntitiesByName](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) メソッド|Exchange Server でアイテム本体上のよく知られたエンティティを抽出する際の制限。 Exchange Server では、その制限を超えるエンティティが無視されます。 この制限は、アドインがルールを使用するかどうかには依存し `ItemHasKnownEntity` ません。|
|Exchange Web サービス|文字数 1 MB|[mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) メソッド|通話に対する要求または応答の制限 `Mailbox.makeEwsRequestAsync` 。|
|受信者|100 の受信者|[item.requiredAttendees](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) プロパティ<br/> <br/>[item.optionalAttendees](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) プロパティ<br/> <br/>[item.to](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) プロパティ<br/> <br/>[item.cc](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) プロパティ<br/> <br/>[Recipients.addAsync](/javascript/api/outlook/office.Recipients#addasync-recipients--options--callback-) メソッド<br/> <br/>[Recipient.getAsync](/javascript/api/outlook/office.Recipients#getasync-options--callback-) メソッド<br/> <br/>[Recipient.setAsync](/javascript/api/outlook/office.Recipients#setasync-recipients--options--callback-) メソッド|各プロパティで指定された受信者の制限|
|表示名|255 文字|[EmailAddressDetails.displayName](/javascript/api/outlook/office.emailaddressdetails#displayname) プロパティ<br/><br/> [Recipients](/javascript/api/outlook/office.Recipients) オブジェクト<br/><br/> `item.requiredAttendees` プロパティ<br/><br/> `item.optionalAttendees` プロパティ <br/><br/>`item.to` プロパティ <br/><br/>`item.cc` プロパティ|予定やメッセージの表示名の長さの制限。|
|件名の設定|255 文字|[Mailbox.displayNewAppointmentForm](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) メソッド<br/><br/> [Subject.setAsync](/javascript/api/outlook/office.Subject#setasync-subject--options--callback-) メソッド|新しい予定のフォームの件名、または予定やメッセージの件名の設定に関する制限。|
|場所の設定|255 文字|[Location.setAsync](/javascript/api/outlook/office.Location#setasync-location--options--callback-) メソッド|予定または会議出席依頼の場所の設定に関する制限。|
|新しい予定のフォームの本文|文字数 32 KB|`Mailbox.displayNewAppointmentForm` 手段|新しい予定のフォームでの本文に関する制限。|
|既存のアイテムの本文の表示|文字数 32 KB|[mailbox.displayAppointmentForm](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) メソッド<br/><br/> [mailbox.displayMessageForm](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) メソッド|Web 上の Outlook およびモバイルデバイスの場合: 既存の予定またはメッセージフォームの本文の制限。|
|本文の設定|文字数 1 MB|[Body.prependAsync](/javascript/api/outlook/office.Body#prependasync-data--options--callback-) メソッド<br/> <br/>[Body.setAsync](/javascript/api/outlook/office.Body#setasync-data--options--callback-)<br/><br/>[Body.setSelectedDataAsync](/javascript/api/outlook/office.Body#setselecteddataasync-data--options--callback-) メソッド|予定またはメッセージ アイテムの本文の設定に関する制限。|
|添付ファイルの数|Outlook on the web およびモバイルデバイスの499ファイル|[item.addFileAttachmentAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) メソッド|アイテムを送信するために添付できるファイル数の制限。 Web 上の Outlook とモバイルデバイスでは、通常、ユーザーインターフェイスとと共に499ファイルへの添付が制限されて `addFileAttachmentAsync` います。 Outlook リッチ クライアントでは、添付ファイルの数は特に制限されません。 ただし、すべての Outlook クライアントで、ユーザーの Exchange サーバーが構成されている添付ファイルのサイズの制限を監視します。 "添付ファイルのサイズ" については、次の行を参照してください。|
|添付ファイルのサイズ|Exchange Server に依存|`item.addFileAttachmentAsync` 手段|管理者がユーザーのメールボックスの Exchange Server で構成できるアイテムの、すべての添付ファイルのサイズには制限があります。Outlook リッチ クライアントの場合、アイテムの添付ファイルの数が制限されます。 Web およびモバイルデバイスの Outlook では、2つの制限のうち、添付ファイルの数とすべての添付ファイルのサイズのうち、アイテムの実際の添付ファイルを制限します。|
|添付ファイルの名前|255 文字|`item.addFileAttachmentAsync` 手段|アイテムに追加する添付ファイルのファイル名の長さを制限します。|
|添付ファイルの URI|2048 文字|`item.addFileAttachmentAsync` 手段|アイテムに添付ファイルとして追加するファイル名の URI に関する制限。|
|添付ファイルの ID|100 文字|[item.addItemAttachmentAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) メソッド<br/><br/> [item.removeAttachmentAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) メソッド|アイテムに追加またはアイテムから削除する添付ファイルの ID の長さに関する制限。|
|非同期呼び出し|3 つの呼び出し|`item.addFileAttachmentAsync` 手段<br/><br/>`item.addItemAttachmentAsync` 手段<br/><br/><br/>`item.removeAttachmentAsync` 手段<br/><br/> [Body.getTypeAsync](/javascript/api/outlook/office.Body#gettypeasync-options--callback-) メソッド<br/><br/>`Body.prependAsync` 手段<br/><br/>`Body.setSelectedDataAsync` 手段<br/><br/> [CustomProperties.saveAsync](/javascript/api/outlook/office.CustomProperties#saveasync-callback--asynccontext-) メソッド<br/><br/><br/> [item.LoadCustomPropertiesAysnc](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) メソッド<br/><br/><br/> [Location.getAsync](/javascript/api/outlook/office.Location#getasync-options--callback-) メソッド<br/><br/>`Location.setAsync` 手段<br/><br/> [mailbox.getCallbackTokenAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) メソッド<br/><br/> [mailbox.getUserIdentityTokenAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) メソッド<br/><br/> [mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) メソッド<br/><br/>`Recipients.addAsync` 手段<br/><br/> [Recipients.getAsync](/javascript/api/outlook/office.Recipients#getasync-options--callback-) メソッド<br/><br/>`Recipients.setAsync` 手段<br/><br/> [RoamingSettings.saveAsync](/javascript/api/outlook/office.RoamingSettings#saveasync-callback-) メソッド<br/><br/> [Subject.getAsync](/javascript/api/outlook/office.Subject#getasync-options--callback-) メソッド<br/><br/>`Subject.setAsync` 手段<br/><br/> [Time.getAsync](/javascript/api/outlook/office.Time#getasync-options--callback-) メソッド<br/><br/> [Time.setAsync](/javascript/api/outlook/office.Time#setasync-datetime--options--callback-) メソッド|Web またはモバイルデバイスの場合: 同時非同期呼び出し数の制限は、サーバーへの非同期呼び出しの数が限られている場合に限り、一度に同時に行うことができます。 |

## <a name="see-also"></a>関連項目

- [テスト用に Outlook アドインを展開してインストールする](testing-and-tips.md)
- [Outlook アドインに関するプライバシー、アクセス許可、セキュリティ](../concepts/privacy-and-security.md)
