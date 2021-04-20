---
title: Outlook アドインで代理人アクセスのシナリオを有効にする
description: 代理人アクセスについて簡単に説明し、アドインサポートを構成する方法について説明します。
ms.date: 02/09/2021
localization_priority: Normal
ms.openlocfilehash: 598f931dbf3a4be8adf029838084ec0767bf6518
ms.sourcegitcommit: fefc279b85e37463413b6b0e84c880d9ed5d7ac3
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/12/2021
ms.locfileid: "50234241"
---
# <a name="enable-delegate-access-scenarios-in-an-outlook-add-in"></a>Outlook アドインで代理人アクセスのシナリオを有効にする

メールボックスの所有者は、代理人アクセス機能を使用して、他のユーザーが自分のメールと予定表 [を管理できます](https://support.office.com/article/allow-someone-else-to-manage-your-mail-and-calendar-41c40c04-3bd1-4d22-963a-28eafec25926)。 この記事では、Office JavaScript API がサポートする委任アクセス許可を指定し、Outlook アドインで代理人アクセスシナリオを有効にする方法について説明します。

> [!IMPORTANT]
> Android および iOS 上の Outlook では、代理人アクセスは現在使用できません。 また、この機能は、Outlook on the web の [グループ共有](/microsoft-365/admin/create-groups/compare-groups?view=o365-worldwide&preserve-view=true#shared-mailboxes) メールボックスでは現在使用できません。 この機能は、今後使用可能になる可能性があります。
>
> この機能のサポートは、要件セット 1.8 で導入されました。 この要件セットをサポートする [クライアントおよびプラットフォーム](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) を参照してください。

## <a name="supported-permissions-for-delegate-access"></a>代理人アクセスでサポートされているアクセス許可

次の表では、JavaScript API がサポートする委任Office説明します。

|アクセス許可|値|説明|
|---|---:|---|
|読み取り|1 (000001)|アイテムを読み取り可能。|
|書き込み|2 (000010)|アイテムを作成できます。|
|DeleteOwn|4 (000100)|作成したアイテムのみを削除できます。|
|DeleteAll|8 (001000)|任意のアイテムを削除できます。|
|EditOwn|16 (010000)|作成したアイテムのみを編集できます。|
|EditAll|32 (100000)|任意のアイテムを編集できます。|

> [!NOTE]
> 現在、API は既存の代理人アクセス許可の取得をサポートしていますが、代理人のアクセス許可を設定することはできません。

[DelegatePermissions オブジェクトは](/javascript/api/outlook/office.mailboxenums.delegatepermissions)、代理人のアクセス許可を示すためにビットマスクを使用して実装されます。 ビットマスク内の各位置は特定のアクセス許可を表し、設定されている場合、代理人はそれぞれのアクセス `1` 許可を持っています。 たとえば、右側の 2 番目のビットが次の場合、代理人は書き込 `1` みアクセス **許可を持** つ。 特定のアクセス許可を確認する方法の例については、この記事の[](#perform-an-operation-as-delegate)後半の「代理として操作を実行する」セクションを参照してください。

## <a name="sync-across-mailbox-clients"></a>メールボックス クライアント間での同期

所有者のメールボックスに対する代理人の更新は、通常、メールボックス間で直ちに同期されます。

ただし、REST 操作または Exchange Web サービス (EWS) 操作を使用してアイテムに拡張プロパティを設定した場合、このような変更の同期には数時間かかる場合があります。このような遅延を避けるために [、CustomProperties](/javascript/api/outlook/office.customproperties) オブジェクトと関連する API を使用することをお勧めします。 詳細については、「Outlook アドイン [でメタデータ](metadata-for-an-outlook-add-in.md#custom-data-per-item-in-a-mailbox-custom-properties) を取得および設定する」の「カスタム プロパティ」セクションを参照してください。

> [!IMPORTANT]
> 委任シナリオでは、現在 API によって提供されているトークンと EWS office.jsできません。

## <a name="configure-the-manifest"></a>マニフェストを構成する

アドインで代理人アクセスのシナリオを有効にするには、親要素のマニフェストで [SupportsSharedFolders](../reference/manifest/supportssharedfolders.md) 要素を設定 `true` する必要があります `DesktopFormFactor` 。 現時点では、他のフォーム ファクターはサポートされていません。

代理人からの REST 呼び出しをサポートするには、マニフェスト [のアクセス](../reference/manifest/permissions.md) 許可ノードを次に設定します `ReadWriteMailbox` 。

次の例は、 `SupportsSharedFolders` マニフェストのセクション `true` に設定されている要素を示しています。

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

## <a name="perform-an-operation-as-delegate"></a>デリゲートとして操作を実行する

[item.getSharedPropertiesAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)メソッドを呼び出すことによって、新規作成モードまたは読み取りモードでアイテムの共有プロパティを取得できます。 これにより、代理人のアクセス許可、所有者の電子メール アドレス、REST API のベース URL、およびターゲット メールボックスを現在提供している [SharedProperties](/javascript/api/outlook/office.sharedproperties) オブジェクトが返されます。

次の例は、メッセージまたは予定の共有プロパティを取得し、代理人が **書** き込みアクセス許可を持って、REST 呼び出しを行う方法を示しています。

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
> 代理人は、REST を使用して、Outlook アイテムまたはグループの投稿に添付された Outlook メッセージのコンテンツ [を取得できます](/graph/outlook-get-mime-message#get-mime-content-of-an-outlook-message-attached-to-an-outlook-item-or-group-post)。

## <a name="handle-calling-rest-on-shared-and-non-shared-items"></a>共有アイテムと非共有アイテムに対する REST の呼び出しの処理

アイテムが共有されるかどうかに関して、アイテムに対して REST 操作を呼び出す場合は、API を使用してアイテムが共有 `getSharedPropertiesAsync` されているかどうかを判断できます。 その後、適切なオブジェクトを使用して操作の REST URL を作成できます。

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

アドインのシナリオに応じて、代理人の状況を処理する際に考慮する必要がある制限があります。

### <a name="rest-and-ews"></a>REST と EWS

アドインは REST を使用できますが、EWS は使用できません。また、所有者のメールボックスへの REST アクセスを有効にするには、アドインのアクセス許可を設定 `ReadWriteMailbox` する必要があります。

### <a name="message-compose-mode"></a>メッセージ作成モード

メッセージ作成モードでは、次の条件が満たされない限り、Outlook on the web または Windows では [getSharedPropertiesAsync](/javascript/api/outlook/office.messagecompose#getsharedpropertiesasync-options--callback-) はサポートされません。

1. 所有者は、代理人と少なくとも 1 つのメールボックス フォルダーを共有します。
1. 代理人は、共有フォルダー内のメッセージを下書きします。

    例:

    - 代理人は、共有フォルダー内の電子メールに返信または転送します。
    - 代理人は下書きメッセージを保存し、それを自分の **下書き** フォルダーから共有フォルダーに移動します。 代理人は、共有フォルダーから下書きを開き、作成を続行します。

メッセージが送信された後は、通常、代理人の [送信されたアイテム] **フォルダーにあります** 。

## <a name="see-also"></a>関連項目

- [自分のメールと予定表の管理を他のユーザーに許可する](https://support.office.com/article/allow-someone-else-to-manage-your-mail-and-calendar-41c40c04-3bd1-4d22-963a-28eafec25926)
- [Microsoft 365 での予定表の共有](https://support.office.com/article/calendar-sharing-in-office-365-b576ecc3-0945-4d75-85f1-5efafb8a37b4)
- [マニフェスト要素を順序付けする方法](../develop/manifest-element-ordering.md)
- [マスク (コンピューティング)](https://en.wikipedia.org/wiki/Mask_(computing))
- [JavaScript のビット演算子](https://www.w3schools.com/js/js_bitwise.asp)