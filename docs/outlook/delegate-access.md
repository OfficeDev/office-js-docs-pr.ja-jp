---
title: Outlook アドインで共有フォルダーと共有メールボックスのシナリオを有効にする
description: 共有フォルダー (a.k.a) のアドイン サポートを構成する方法について説明します。 委任アクセス) と共有メールボックス。
ms.date: 09/12/2022
ms.localizationpriority: medium
ms.openlocfilehash: 1c6884c18e4cb9916fcec20e6b732b0d20918e2f
ms.sourcegitcommit: 54a7dc07e5f31dd5111e4efee3e85b4643c4bef5
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/21/2022
ms.locfileid: "67857558"
---
# <a name="enable-shared-folders-and-shared-mailbox-scenarios-in-an-outlook-add-in"></a>Outlook アドインで共有フォルダーと共有メールボックスのシナリオを有効にする

この記事では、Office JavaScript API がサポートするアクセス許可を含め、Outlook アドインで共有フォルダー (委任アクセスとも呼ばれます) と共有メールボックス (プレビュー [段階](/javascript/api/requirement-sets/outlook/preview-requirement-set/outlook-requirement-set-preview#shared-mailboxes)) のシナリオを有効にする方法について説明します。

## <a name="supported-clients-and-platforms"></a>サポートされているクライアントとプラットフォーム

次の表は、この機能でサポートされているクライアントとサーバーの組み合わせを示しています。これには、必要最小限の累積更新プログラム (該当する場合) が含まれます。 除外された組み合わせはサポートされていません。

| Client | Exchange Online | Exchange 2019 オンプレミス<br>(累積的な更新プログラム 1 以降) | Exchange 2016 オンプレミス<br>(累積的な更新プログラム 6 以降) | Exchange 2013 オンプレミス |
|---|:---:|:---:|:---:|:---:|
|Windows:<br>バージョン 1910 (ビルド 12130.20272) 以降|はい|はい\*|はい\*|はい\*|
|Mac：<br>ビルド 16.47 以降|はい|はい|はい|はい|
|Web ブラウザー:<br>最新の Outlook UI|あり|該当なし|該当なし|該当なし|
|Web ブラウザー:<br>クラシック Outlook UI|該当なし|いいえ|いいえ|いいえ|

> [!NOTE]
> \* オンプレミス Exchange 環境でのこの機能のサポートは、現在のチャネルのバージョン 2206 (ビルド 15330.20000) と月次エンタープライズ チャネルのバージョン 2207 (ビルド 15427.20000) 以降で利用できます。

> [!IMPORTANT]
> この機能のサポートは [、要件セット 1.8](/javascript/api/requirement-sets/outlook/requirement-set-1.8/outlook-requirement-set-1.8) で導入されました (詳細については、 [クライアントとプラットフォーム](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets#requirement-sets-supported-by-exchange-servers-and-outlook-clients)を参照してください)。 ただし、この機能のサポート マトリックスは要件セットのスーパーセットであることに注意してください。

## <a name="supported-setups"></a>サポートされているセットアップ

次のセクションでは、共有メールボックス (プレビュー段階) と共有フォルダーでサポートされている構成について説明します。 機能 API は、他の構成では期待どおりに機能しない可能性があります。 構成する方法を学習するプラットフォームを選択します。

### <a name="windows"></a>[Windows](#tab/windows)

#### <a name="shared-folders"></a>共有フォルダー

メールボックスの所有者は、最初 [に代理人へのアクセスを提供する](https://support.microsoft.com/office/41c40c04-3bd1-4d22-963a-28eafec25926)必要があります。 代理人は、「別のユーザーの [メールアイテムと予定表アイテムを管理](https://support.microsoft.com/office/afb79d6b-2967-43b9-a944-a6b953190af5)する」の「別のユーザーのメールボックスをプロファイルに追加する」セクションに記載されている手順に従う必要があります。

#### <a name="shared-mailboxes-preview"></a>共有メールボックス (プレビュー)

Exchange サーバー管理者は、アクセスする一連のユーザーの共有メールボックスを作成および管理できます。 [Exchange Online](/exchange/collaboration-exo/shared-mailboxes)および[オンプレミスの Exchange 環境](/exchange/collaboration/shared-mailboxes/create-shared-mailboxes)がサポートされています。

"automapping" と呼ばれるExchange Server機能は既定でオンになっています。つまり、Outlook を閉じて再度開いた後、その後、共有メールボックスがユーザーの Outlook アプリ[に自動的に表示](/microsoft-365/admin/email/create-a-shared-mailbox?view=o365-worldwide&preserve-view=true#add-the-shared-mailbox-to-outlook)されます。 ただし、管理者が自動マッピングをオフにした場合は、「Outlook で共有メールボックスを [開いて使用](https://support.microsoft.com/office/d94a8e9e-21f1-4240-808b-de9c9c088afd)する」の「Outlook に共有メールボックスを追加する」セクションで説明されている手動の手順に従う必要があります。

> [!WARNING]
> パスワードを使用して共有メールボックスにサインイン **しないでください** 。 この場合、機能 API は機能しません。

### <a name="web-browser---modern-outlook"></a>[Web ブラウザー - モダン Outlook](#tab/modern)

#### <a name="shared-folders"></a>共有フォルダー

メールボックス所有者は、最初にメールボックス フォルダー [のアクセス許可を更新して代理人へのアクセスを提供](https://www.microsoft.com/microsoft-365/blog/2013/09/04/configuring-delegate-access-in-outlook-web-app/) する必要があります。 代理人は、「別のユーザーのメールボックスにアクセスする」の「Outlook Web Appで別のユーザーのメールボックスをフォルダーの一覧に追加する」セクションに記載されている手順[に従う](https://support.microsoft.com/office/a909ad30-e413-40b5-a487-0ea70b763081)必要があります。

#### <a name="shared-mailboxes"></a>共有メールボックス

Outlook アドインの共有メールボックスシナリオは、現在、最新のOutlook on the webではサポートされていません。

### <a name="mac"></a>[Mac](#tab/unix)

#### <a name="shared-mailboxes-preview"></a>共有メールボックス (プレビュー)

メールと予定表は、代理人または共有メールボックス ユーザーと共有されます。 メッセージモードと予定読み取りモードと作成モードでは、代理人またはユーザーがアドインを使用できます。

#### <a name="shared-folders"></a>共有フォルダー

**受信トレイ** フォルダーが代理人と共有されている場合は、メッセージ読み取りモードで代理人がアドインを使用できます。

**下書き** フォルダーもデリゲートと共有されている場合、アドインは作成モードで使用できます。

#### <a name="local-shared-calendar-new-model"></a>ローカル共有予定表 (新しいモデル)

予定表の所有者が代理人と予定表を明示的に共有している場合 (メールボックス全体を共有できない場合があります)、アドインは、予定の読み取りモードと作成モードで代理人が使用できます。

#### <a name="remote-shared-calendar-previous-model"></a>リモート共有予定表 (以前のモデル)

予定表の所有者が予定表への広範なアクセスを許可した場合 (たとえば、特定の DL または組織全体に編集可能にするなど)、ユーザーは間接的または暗黙的なアクセス許可を持つ可能性があり、アドインは予定の読み取りモードと作成モードのユーザーが使用できます。

---

アドインが一般的にアクティブ化する場所とアクティブ化しない場所の詳細については、Outlook アドインの概要ページの「 [アドインで使用できるメールボックス アイテム](outlook-add-ins-overview.md#mailbox-items-available-to-add-ins) 」セクションを参照してください。

## <a name="supported-permissions"></a>サポートされているアクセス許可

次の表では、Office JavaScript API が代理人および共有メールボックス ユーザーに対してサポートするアクセス許可について説明します。

|アクセス許可|値|説明|
|---|---:|---|
|読み取り|1 (000001)|アイテムを読み取ることができます。|
|書き込み|2 (000010)|アイテムを作成できます。|
|DeleteOwn|4 (000100)|削除できるのは、作成したアイテムのみです。|
|DeleteAll|8 (001000)|任意のアイテムを削除できます。|
|EditOwn|16 (010000)|作成したアイテムのみを編集できます。|
|EditAll|32 (100000)|任意のアイテムを編集できます。|

> [!NOTE]
> 現在、API は既存のアクセス許可の取得をサポートしていますが、アクセス許可の設定はサポートされていません。

[DelegatePermissions](/javascript/api/outlook/office.mailboxenums.delegatepermissions) オブジェクトは、アクセス許可を示すためにビットマスクを使用して実装されます。 ビットマスク内の各位置は特定のアクセス許可を表し、それが設定 `1` されている場合、ユーザーはそれぞれのアクセス許可を持ちます。 たとえば、右側の 2 番目のビットが次のビットの場合、 `1`ユーザーは **書き込み** アクセス許可を持ちます。 この記事の後半の「 [委任メールボックスまたは共有メールボックスユーザーとして操作を実行する](#perform-an-operation-as-delegate-or-shared-mailbox-user) 」セクションで、特定のアクセス許可を確認する方法の例を確認できます。

## <a name="sync-across-shared-folder-clients"></a>共有フォルダー クライアント間で同期する

所有者のメールボックスに対する代理人の更新は、通常、メールボックス間ですぐに同期されます。

ただし、REST または Exchange Web Services (EWS) 操作を使用してアイテムに拡張プロパティを設定した場合、このような変更の同期には数時間かかる可能性があります。このような遅延を回避するには、代わりに [CustomProperties](/javascript/api/outlook/office.customproperties) オブジェクトと関連する API を使用することをお勧めします。 詳細については、「Outlook アドインでメタデータを取得して設定する」の記事の [カスタム プロパティセクション](metadata-for-an-outlook-add-in.md#custom-data-per-item-in-a-mailbox-custom-properties) を参照してください。

> [!IMPORTANT]
> デリゲート シナリオでは、office.js API によって現在提供されているトークンで EWS を使用することはできません。

## <a name="configure-the-manifest"></a>マニフェストを構成する

アドインで共有フォルダーと共有メールボックスのシナリオを有効にするには、親`DesktopFormFactor`要素の下のマニフェストで [SupportsSharedFolders](/javascript/api/manifest/supportssharedfolders) 要素を設定する`true`必要があります。 現時点では、他のフォーム ファクターはサポートされていません。

デリゲートからの REST 呼び出しをサポートするには、マニフェスト`ReadWriteMailbox`の [[アクセス許可]](/javascript/api/manifest/permissions) ノードを [ .

次の例は、マニフェストの `SupportsSharedFolders` セクションに設定 `true` されている要素を示しています。

```XML
...
<VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
    ...
    <Hosts>
      <Host xsi:type="MailHost">
        <DesktopFormFactor>
          <SupportsSharedFolders>true</SupportsSharedFolders>
          <FunctionFile resid="residDesktopFuncUrl" />
          <ExtensionPoint xsi:type="MessageReadCommandSurface">
            <!-- configure selected extension point -->
          </ExtensionPoint>

          <!-- You can define more than one ExtensionPoint element as needed -->

        </DesktopFormFactor>
      </Host>
    </Hosts>
    ...
  </VersionOverrides>
</VersionOverrides>
...
```

## <a name="perform-an-operation-as-delegate-or-shared-mailbox-user"></a>代理人または共有メールボックス ユーザーとして操作を実行する

[item.getSharedPropertiesAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods) メソッドを呼び出すことで、作成モードまたは読み取りモードでアイテムの共有プロパティを取得できます。 これにより、現在ユーザーのアクセス許可、所有者の電子メール アドレス、REST API のベース URL、およびターゲット メールボックスを提供する [SharedProperties](/javascript/api/outlook/office.sharedproperties) オブジェクトが返されます。

次の例は、メッセージまたは予定の共有プロパティを取得し、代理人または共有メールボックスユーザーが **書き込み** アクセス許可を持っているかどうかを確認し、REST 呼び出しを行う方法を示しています。

```js
function performOperation() {
  Office.context.mailbox.getCallbackTokenAsync({
      isRest: true
    },
    function (asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded && asyncResult.value !== "") {
        Office.context.mailbox.item.getSharedPropertiesAsync({
            // Pass auth token along.
            asyncContext: asyncResult.value
          },
          function (asyncResult1) {
            let sharedProperties = asyncResult1.value;
            let delegatePermissions = sharedProperties.delegatePermissions;

            // Determine if user can do the expected operation.
            // E.g., do they have Write permission?
            if ((delegatePermissions & Office.MailboxEnums.DelegatePermissions.Write) != 0) {
              // Construct REST URL for your operation.
              // Update <version> placeholder with actual Outlook REST API version e.g. "v2.0".
              // Update <operation> placeholder with actual operation.
              let rest_url = sharedProperties.targetRestUrl + "/<version>/users/" + sharedProperties.targetMailbox + "/<operation>";
  
              $.ajax({
                  url: rest_url,
                  dataType: 'json',
                  headers:
                  {
                    "Authorization": "Bearer " + asyncResult1.asyncContext
                  }
                }
              ).done(
                function (response) {
                  console.log("success");
                }
              ).fail(
                function (error) {
                  console.log("error message");
                }
              );
            }
          }
        );
      }
    }
  );
}
```

> [!TIP]
> 代理人として、REST を使用して [、Outlook アイテムまたはグループ投稿に添付されている Outlook メッセージのコンテンツを取得](/graph/outlook-get-mime-message#get-mime-content-of-an-outlook-message-attached-to-an-outlook-item-or-group-post)できます。

## <a name="handle-calling-rest-on-shared-and-non-shared-items"></a>共有アイテムと非共有アイテムに対する REST の呼び出しを処理する

アイテムが共有されているかどうかに関係なく、アイテムに対して REST 操作を呼び出す場合は、API を `getSharedPropertiesAsync` 使用してアイテムが共有されているかどうかを判断できます。 その後、適切なオブジェクトを使用して操作の REST URL を作成できます。

```js
if (item.getSharedPropertiesAsync) {
  // In Windows, Mac, and the web client, this indicates a shared item so use SharedProperties properties to construct the REST URL.
  // Add-ins don't activate on shared items in mobile so no need to handle.

  // Perform operation for shared item.
} else {
  // In general, this is not a shared item, so construct the REST URL using info from the Call REST APIs article:
  // https://learn.microsoft.com/office/dev/add-ins/outlook/use-rest-api

  // Perform operation for non-shared item.
}
```

## <a name="limitations"></a>制限事項

アドインのシナリオによっては、共有フォルダーまたは共有メールボックスの状況を処理する際に考慮すべき制限事項がいくつかあります。

### <a name="message-compose-mode"></a>メッセージ作成モード

メッセージ作成モードでは、次の条件が満たされない限り、Outlook on the webまたは Windows では [getSharedPropertiesAsync](/javascript/api/outlook/office.messagecompose#outlook-office-messagecompose-getsharedpropertiesasync-member(1)) はサポートされません。

a. **アクセスの委任/共有フォルダー**

1. メールボックスの所有者がメッセージを開始します。 これは、新しいメッセージ、応答、または転送にすることができます。
1. メッセージを保存し、自分の **下書き** フォルダーから代理人と共有されているフォルダーに移動します。
1. 代理人は、共有フォルダーから下書きを開き、作成を続行します。

b. **共有メールボックス (Outlook on Windows にのみ適用)**

1. 共有メールボックス ユーザーがメッセージを開始します。 これは、新しいメッセージ、応答、または転送にすることができます。
1. メッセージを保存し、自分の **下書き** フォルダーから共有メールボックス内のフォルダーに移動します。
1. 別の共有メールボックス ユーザーが共有メールボックスから下書きを開き、作成を続行します。

メッセージが共有コンテキストに入り、これらの共有シナリオをサポートするアドインでアイテムの共有プロパティを取得できるようになりました。 メッセージが送信された後、通常は送信者の **[送信済みアイテム]** フォルダーにあります。

### <a name="rest-and-ews"></a>REST と EWS

アドインは REST を使用でき、所有者のメールボックスまたは共有メールボックスへの REST アクセスを有効にするには `ReadWriteMailbox` 、アドインのアクセス許可を設定する必要があります(該当する場合)。 EWS はサポートされていません。

### <a name="user-or-shared-mailbox-hidden-from-an-address-list"></a>アドレス一覧から非表示になっているユーザーまたは共有メールボックス

管理者がグローバル アドレス一覧 (GAL) などのアドレス一覧からユーザーまたは共有メールボックス アドレスを隠した場合、メールボックス レポートで開かれた影響を受けるメール アイテムは null として表示 `Office.context.mailbox.item` されます。 たとえば、ユーザーが GAL から非表示になっている共有メールボックスでメール アイテムを開いた場合、 `Office.context.mailbox.item` そのメール アイテムは null であることを表します。

## <a name="see-also"></a>関連項目

- [自分のメールと予定表の管理を他のユーザーに許可する](https://support.microsoft.com/office/41c40c04-3bd1-4d22-963a-28eafec25926)
- [Microsoft 365 での予定表共有](https://support.microsoft.com/office/b576ecc3-0945-4d75-85f1-5efafb8a37b4)
- [Outlook に共有メールボックスを追加する](/microsoft-365/admin/email/create-a-shared-mailbox?view=o365-worldwide&preserve-view=true#add-the-shared-mailbox-to-outlook)
- [マニフェスト要素を並べ替える方法](../develop/manifest-element-ordering.md)
- [Mask (コンピューティング)](https://en.wikipedia.org/wiki/Mask_(computing))
- [JavaScript ビットごとの演算子](https://www.w3schools.com/js/js_bitwise.asp)
