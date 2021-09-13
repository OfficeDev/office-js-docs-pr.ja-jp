---
title: 共有フォルダーと共有メールボックスのシナリオを、Outlookアドインで有効にする
description: 共有フォルダー (a.k.a) のアドイン サポートを構成する方法について説明します。 アクセスを委任する) と共有メールボックスを使用します。
ms.date: 07/02/2021
ms.localizationpriority: medium
ms.openlocfilehash: 65850699612e9dc48dfe7cc1aed5b00ce5b79012
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/12/2021
ms.locfileid: "59154174"
---
# <a name="enable-shared-folders-and-shared-mailbox-scenarios-in-an-outlook-add-in"></a>共有フォルダーと共有メールボックスのシナリオを、Outlookアドインで有効にする

この記事では、Outlook アドインで共有フォルダー (代理人アクセスとも呼ばれる) と共有メールボックス (プレビュー[中)](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md#shared-mailboxes)のシナリオ (Office JavaScript API がサポートするアクセス許可を含む) を有効にする方法について説明します。

> [!IMPORTANT]
> この機能のサポートは、要件セット [1.8 で導入されました](../reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md)。 この要件セットをサポートする [クライアントおよびプラットフォーム](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) を参照してください。

## <a name="supported-setups"></a>サポートされているセットアップ

次のセクションでは、共有メールボックス (プレビュー中) と共有フォルダーでサポートされる構成について説明します。 機能 API は、他の構成では期待通り動作しない場合があります。 構成方法を学習するプラットフォームを選択します。

### <a name="windows"></a>[Windows](#tab/windows)

#### <a name="shared-folders"></a>共有フォルダー

メールボックスの所有者は、まず代理人 [へのアクセス権を提供する必要があります](https://support.microsoft.com/office/41c40c04-3bd1-4d22-963a-28eafec25926)。 代理人は、「他のユーザーのメールと予定表アイテムを管理する」の「自分のプロファイルに他のユーザーのメールボックスを追加する」セクションに記載されている手順に従う [必要があります](https://support.microsoft.com/office/afb79d6b-2967-43b9-a944-a6b953190af5)。

#### <a name="shared-mailboxes-preview"></a>共有メールボックス (プレビュー)

Exchange管理者は、アクセスするユーザーのセットの共有メールボックスを作成および管理できます。 現時点では[、Exchange Online](/exchange/collaboration-exo/shared-mailboxes)この機能でサポートされている唯一のサーバー バージョンです。

"自動Exchange Server" と呼ばれる機能が既定でオンになっています。つまり、Outlook が閉[](/microsoft-365/admin/email/create-a-shared-mailbox?view=o365-worldwide&preserve-view=true#add-the-shared-mailbox-to-outlook)じて再び開いた後、共有メールボックスはユーザーの Outlook アプリに自動的に表示されます。 ただし、管理者が自動マップをオフにした場合、ユーザーは記事「開く」の「Outlook に共有メールボックスを追加して、Outlook で共有メールボックスを使用する」に示されている手動の手順[に従う必要があります](https://support.microsoft.com/office/d94a8e9e-21f1-4240-808b-de9c9c088afd)。

> [!WARNING]
> パスワード **を** 使用して共有メールボックスにサインインしない。 この場合、機能 API は機能しません。

### <a name="web-browser---modern-outlook"></a>[Web ブラウザー - モダン Outlook](#tab/modern)

#### <a name="shared-folders"></a>共有フォルダー

メールボックスの所有者は、まずメールボックス フォルダー [のアクセス許可を](https://www.microsoft.com/microsoft-365/blog/2013/09/04/configuring-delegate-access-in-outlook-web-app/) 更新して代理人へのアクセスを提供する必要があります。 代理人は、記事「他のユーザーのメールボックスにアクセスする」の「Outlook Web App のフォルダー リストに他のユーザーのメールボックスを追加する」に記載されている手順に従う[必要があります](https://support.microsoft.com/office/a909ad30-e413-40b5-a487-0ea70b763081)。

#### <a name="shared-mailboxes-preview"></a>共有メールボックス (プレビュー)

Exchange管理者は、アクセスするユーザーのセットの共有メールボックスを作成および管理できます。 現時点では[、Exchange Online](/exchange/collaboration-exo/shared-mailboxes)この機能でサポートされている唯一のサーバー バージョンです。

アクセスを受け取った後、共有メールボックス ユーザーは、「Open」の「共有メールボックスを追加してプライマリ メールボックスの下に表示する」セクションに示されている手順に従って、Outlook on the web で共有メールボックスを使用する[必要があります](https://support.microsoft.com/office/98b5a90d-4e38-415d-a030-f09a4cd28207)。

> [!WARNING]
> " **別の** メールボックスを開く" などの他のオプションを使用しない。 機能 API が正しく動作しない場合があります。

---

アドインが一般的にアクティブ化する場所とアクティブ化しない場所の詳細については、「Outlook[](outlook-add-ins-overview.md#mailbox-items-available-to-add-ins)アドインの概要」ページの「アドインで使用可能なメールボックス アイテム」セクションを参照してください。

## <a name="supported-permissions"></a>サポートされているアクセス許可

次の表では、代理人および共有メールボックス ユーザー Office JavaScript API でサポートされるアクセス許可について説明します。

|アクセス許可|値|説明|
|---|---:|---|
|読み取り|1 (000001)|アイテムを読み取り可能。|
|書き込み|2 (000010)|アイテムを作成できます。|
|DeleteOwn|4 (000100)|作成したアイテムのみを削除できます。|
|DeleteAll|8 (001000)|任意のアイテムを削除できます。|
|EditOwn|16 (010000)|作成したアイテムのみを編集できます。|
|EditAll|32 (100000)|任意のアイテムを編集できます。|

> [!NOTE]
> 現在、API は既存のアクセス許可の取得をサポートしていますが、アクセス許可を設定することはできません。

[DelegatePermissions オブジェクト](/javascript/api/outlook/office.mailboxenums.delegatepermissions)は、アクセス許可を示すためにビットマスクを使用して実装されます。 ビットマスク内の各位置は、特定のアクセス許可を表し、それが設定されている場合、 `1` ユーザーはそれぞれのアクセス許可を持っています。 たとえば、右側の 2 番目のビットが次の場合 `1` 、ユーザーは書き込みアクセス **許可を持** つ。 特定のアクセス許可を確認する方法の例については、後の「[](#perform-an-operation-as-delegate-or-shared-mailbox-user)代理人または共有メールボックス ユーザーとして操作を実行する」セクションを参照してください。

## <a name="sync-across-shared-folder-clients"></a>共有フォルダー クライアント間の同期

所有者のメールボックスに対する代理人の更新は、通常、すぐにメールボックス間で同期されます。

ただし、REST または Exchange Web サービス (EWS) 操作を使用してアイテムに拡張プロパティを設定した場合、このような変更は同期に数時間かかる可能性があります。このような遅延を回避するには[、CustomProperties](/javascript/api/outlook/office.customproperties)オブジェクトと関連する API を使用することをお勧めします。 詳細については、「Get [and](metadata-for-an-outlook-add-in.md#custom-data-per-item-in-a-mailbox-custom-properties) set metadata in the Outlook」の記事を参照してください。

> [!IMPORTANT]
> 代理人シナリオでは、API によって現在提供されているトークンで EWS をoffice.jsできません。

## <a name="configure-the-manifest"></a>マニフェストを構成する

アドインで共有フォルダーと共有メールボックスのシナリオを有効にするには、親要素の下のマニフェストで [SupportsSharedFolders](../reference/manifest/supportssharedfolders.md) 要素を設定 `true` する必要があります `DesktopFormFactor` 。 現時点では、他のフォーム ファクターはサポートされていません。

代理人からの REST 呼び出しをサポートするには、マニフェストの [Permissions](../reference/manifest/permissions.md) ノードをに設定します `ReadWriteMailbox` 。

次の例は、 `SupportsSharedFolders` マニフェストのセクションに `true` 設定された要素を示しています。

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

[item.getSharedPropertiesAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)メソッドを呼び出すことによって、作成モードまたは読み取りモードでアイテムの共有プロパティを取得できます。 これにより、現在ユーザーのアクセス許可、所有者の電子メール アドレス、REST API の基本 URL、およびターゲット メールボックスを提供する [SharedProperties](/javascript/api/outlook/office.sharedproperties) オブジェクトが返されます。

次の例は、メッセージまたは予定の共有プロパティを取得し、代理人または共有メールボックス ユーザーが **書** き込みアクセス許可を持つか確認し、REST 呼び出しを行う方法を示しています。

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
> 代理人として REST を使用して、アイテムまたはグループの投稿にOutlookメッセージのOutlook[を取得できます](/graph/outlook-get-mime-message#get-mime-content-of-an-outlook-message-attached-to-an-outlook-item-or-group-post)。

## <a name="handle-calling-rest-on-shared-and-non-shared-items"></a>共有アイテムと非共有アイテムの REST の呼び出しを処理する

アイテムに対して REST 操作を呼び出す場合は、アイテムが共有されるかどうかに関して、API を使用してアイテムが共有 `getSharedPropertiesAsync` されているかどうかを判断できます。 その後、適切なオブジェクトを使用して操作の REST URL を作成できます。

```js
if (item.getSharedPropertiesAsync) {
  // In Windows, Mac, and the web client, this indicates a shared item so use SharedProperties properties to construct the REST URL.
  // Add-ins don't activate on shared items in mobile so no need to handle.

  // Perform operation for shared item.
} else {
  // In general, this is not a shared item, so construct the REST URL using info from the Call REST APIs article:
  // https://docs.microsoft.com/office/dev/add-ins/outlook/use-rest-api

  // Perform operation for non-shared item.
}
```

## <a name="limitations"></a>制限事項

アドインのシナリオによっては、共有フォルダーや共有メールボックスの状況を処理する際に考慮すべきいくつかの制限があります。

### <a name="message-compose-mode"></a>メッセージ作成モード

メッセージ作成モードでは、次の条件を満たしていない限り、Outlook on the web または Windows で[getSharedPropertiesAsync](/javascript/api/outlook/office.messagecompose#getSharedPropertiesAsync_options__callback_)はサポートされません。

a. **アクセス/共有フォルダーの委任**

1. メールボックスの所有者がメッセージを開始します。 これは、新しいメッセージ、返信、または転送を指定できます。
1. メッセージを保存し、そのメッセージを自分の **下** 書きフォルダーから代理人と共有するフォルダーに移動します。
1. 代理人は、共有フォルダーから下書きを開き、作成を続行します。

b. **共有メールボックス**

1. 共有メールボックス ユーザーがメッセージを開始します。 これは、新しいメッセージ、返信、または転送を指定できます。
1. メッセージを保存し、そのメッセージを自分の **下** 書きフォルダーから共有メールボックス内のフォルダーに移動します。
1. 別の共有メールボックス ユーザーが共有メールボックスから下書きを開き、作成を続行します。

このメッセージは共有コンテキストに追加され、これらの共有シナリオをサポートするアドインはアイテムの共有プロパティを取得できます。 メッセージが送信された後、通常は送信者の [送信アイテム] **フォルダーにあります** 。

### <a name="rest-and-ews"></a>REST と EWS

アドインは REST を使用できます。また、アドインのアクセス許可を設定して、所有者のメールボックスまたは共有メールボックスへの REST アクセスを有効にする必要 `ReadWriteMailbox` があります。 EWS はサポートされていません。

### <a name="user-or-shared-mailbox-hidden-from-an-address-list"></a>アドレス一覧から非表示のユーザーまたは共有メールボックス

管理者がグローバル アドレス一覧 (GAL) などのアドレス一覧からユーザーまたは共有メールボックス のアドレスを非表示にした場合、メールボックス レポートで開いた影響を受けるメール アイテムは `Office.context.mailbox.item` null として開きます。 たとえば、ユーザーが GAL から非表示の共有メールボックスでメール アイテムを開くと、そのメール アイテムは `Office.context.mailbox.item` null になります。

## <a name="see-also"></a>関連項目

- [自分のメールと予定表の管理を他のユーザーに許可する](https://support.microsoft.com/office/41c40c04-3bd1-4d22-963a-28eafec25926)
- [カレンダーの共有Microsoft 365](https://support.microsoft.com/office/b576ecc3-0945-4d75-85f1-5efafb8a37b4)
- [共有メールボックスをユーザーに追加Outlook](/microsoft-365/admin/email/create-a-shared-mailbox?view=o365-worldwide&preserve-view=true#add-the-shared-mailbox-to-outlook)
- [マニフェスト要素を順序付けする方法](../develop/manifest-element-ordering.md)
- [マスク (コンピューティング)](https://en.wikipedia.org/wiki/Mask_(computing))
- [JavaScript ビット演算子](https://www.w3schools.com/js/js_bitwise.asp)