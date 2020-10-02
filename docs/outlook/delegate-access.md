---
title: Outlook アドインで代理人アクセスのシナリオを有効にする
description: 代理人アクセスについて簡単に説明し、アドインサポートを構成する方法について説明します。
ms.date: 09/30/2020
localization_priority: Normal
ms.openlocfilehash: 68e9c8003f8d223a591283fd1a73f0a38bd3c8a4
ms.sourcegitcommit: 6c3a04acde57832feeaaa599148f93af7e3e36ea
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/02/2020
ms.locfileid: "48336420"
---
# <a name="enable-delegate-access-scenarios-in-an-outlook-add-in"></a>Outlook アドインで代理人アクセスのシナリオを有効にする

メールボックスの所有者は代理人アクセス機能を使用して、 [他のユーザーが自分のメールと予定表を管理できるよう](https://support.office.com/article/allow-someone-else-to-manage-your-mail-and-calendar-41c40c04-3bd1-4d22-963a-28eafec25926)にすることができます。 この記事では、Office JavaScript API でサポートされている代理人アクセス許可を指定し、Outlook アドインで代理人アクセスのシナリオを有効にする方法について説明します。

> [!IMPORTANT]
> 代理人アクセスは、Android および iOS の Outlook では現在使用できません。 また、この機能は、web 上の Outlook の [グループ共有メールボックス](/microsoft-365/admin/create-groups/compare-groups?view=o365-worldwide&preserve-view=true#shared-mailboxes) では現在使用できません。 この機能は、今後利用可能になる可能性があります。
>
> この機能のサポートは、要件セット1.8 で導入されました。 この要件セットをサポートする [クライアントおよびプラットフォーム](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) を参照してください。

## <a name="supported-permissions-for-delegate-access"></a>代理人アクセスに対してサポートされるアクセス許可

次の表では、Office JavaScript API でサポートされている代理人アクセス許可について説明します。

|アクセス許可|値|説明|
|---|---:|---|
|Read|1 (000001)|アイテムを読み取ることができます。|
|書き込み|2 (000010)|アイテムを作成できます。|
|DeleteOwn|4 (000100)|は、自分で作成したアイテムのみを削除できます。|
|DeleteAll|8 (001000)|任意のアイテムを削除できます。|
|EditOwn|16 (010000)|は、自分で作成したアイテムのみを編集できます。|
|EditAll|32 (10万)|任意のアイテムを編集できます。|

> [!NOTE]
> 現在、API は既存の代理人アクセス許可の取得をサポートしていますが、代理人アクセス許可は設定しません。

[DelegatePermissions](/javascript/api/outlook/office.mailboxenums.delegatepermissions)オブジェクトは、デリゲートのアクセス許可を示すために、ビットマスクを使用して実装されます。 ビットマスク内の各位置は特定のアクセス許可を表し、それが設定されている場合は `1` 代理人にそれぞれのアクセス許可が付与されます。 たとえば、右側の2番目のビットがの場合、 `1` デリゲートには **書き込み** アクセス許可があります。 この記事で後述する「 [代理人として操作を実行](#perform-an-operation-as-delegate) する」の特定のアクセス許可を確認する方法の例を確認できます。

## <a name="sync-across-mailbox-clients"></a>メールボックスクライアント間での同期

通常、所有者のメールボックスに対する代理人の更新は、メールボックス間で即時に同期されます。

ただし、REST または Exchange Web サービス (EWS) の操作を使用してアイテムの拡張プロパティを設定した場合は、そのような変更を同期するのに数時間かかることがあります。このような遅延を避けるには、 [CustomProperties](/javascript/api/outlook/office.customproperties) オブジェクトと関連する api を使用することをお勧めします。 詳細については、記事「Outlook アドインでメタデータを取得および設定する」の「 [カスタムプロパティ」セクション](metadata-for-an-outlook-add-in.md#custom-data-per-item-in-a-mailbox-custom-properties) を参照してください。

> [!IMPORTANT]
> 代理人のシナリオでは、office.js API によって現在提供されているトークンと共に EWS を使用することはできません。

## <a name="configure-the-manifest"></a>マニフェストを構成する

アドインで代理人アクセスのシナリオを有効にするには、親要素のマニフェスト内の [Supportssharedfolders](../reference/manifest/supportssharedfolders.md) 要素をに設定する必要があり `true` `DesktopFormFactor` ます。 現在、他のフォームファクターはサポートされていません。

代理人からの REST 呼び出しをサポートするには、マニフェストの [ [アクセス許可](../reference/manifest/permissions.md) ] ノードをに設定 `ReadWriteMailbox` します。

次の例は、 `SupportsSharedFolders` マニフェストのセクション内に設定された要素を示して `true` います。

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

## <a name="perform-an-operation-as-delegate"></a>代理人として操作を実行する

アイテムの共有プロパティは、新規作成または閲覧モードで取得できます。そのためには、 [getSharedPropertiesAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) メソッドを呼び出します。 これにより、現在、代理人のアクセス許可、所有者の電子メールアドレス、REST API のベース URL、ターゲットメールボックスを提供する [Sharedproperties](/javascript/api/outlook/office.sharedproperties) オブジェクトが返されます。

次の例は、メッセージまたは予定の共有プロパティを取得する方法、代理人が **書き込み** アクセス許可を持っているかどうかを確認する方法、および REST 呼び出しを行う方法を示しています。

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
> 代理人は、REST を使用して、 [outlook アイテムまたはグループ投稿に添付されている outlook メッセージのコンテンツを取得](/graph/outlook-get-mime-message#get-mime-content-of-an-outlook-message-attached-to-an-outlook-item-or-group-post)できます。

## <a name="handle-calling-rest-on-shared-and-non-shared-items"></a>共有アイテムと共有されていないアイテムの通話を処理する

アイテムが共有されているかどうかにかかわらず、アイテムに対して REST 操作を呼び出す場合は、API を使用して、 `getSharedPropertiesAsync` アイテムが共有されているかどうかを判断できます。 その後、適切なオブジェクトを使用して、操作の REST URL を作成できます。

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

アドインのシナリオに応じて、代理人の状況を処理する際に考慮すべきいくつかの制限があります。

### <a name="rest-and-ews"></a>REST と EWS

アドインでは、EWS ではなく REST を使用でき、アドインのアクセス許可をに設定して、 `ReadWriteMailbox` 所有者のメールボックスへの rest アクセスを有効にする必要があります。

### <a name="message-compose-mode"></a>メッセージ作成モード

メッセージ作成モードでは、次の条件が満たされない限り、 [Getsharedpropertiesasync](/javascript/api/outlook/office.messagecompose#getsharedpropertiesasync-options--callback-) は web 上の Outlook ではサポートされません。

1. 所有者は、代理人と共に少なくとも1つのメールボックスフォルダーを共有します。
1. 代理人は、共有フォルダー内のメッセージを下書きします。

    例:

    - 代理人が共有フォルダー内の電子メールに返信または転送します。
    - 代理人は下書きメッセージを保存し、自分の **下書き** フォルダーから共有フォルダーに移動します。 代理人が共有フォルダーから下書きを開き、作成を続行します。

送信されたメッセージは、通常は代理人の [ **送信済みアイテム** ] フォルダーにあります。

## <a name="see-also"></a>関連項目

- [自分のメールと予定表の管理を他のユーザーに許可する](https://support.office.com/article/allow-someone-else-to-manage-your-mail-and-calendar-41c40c04-3bd1-4d22-963a-28eafec25926)
- [Office365 での予定表の共有](https://support.office.com/article/calendar-sharing-in-office-365-b576ecc3-0945-4d75-85f1-5efafb8a37b4)
- [マニフェスト要素の注文方法](../develop/manifest-element-ordering.md)
- [マスク (コンピューティング)](https://en.wikipedia.org/wiki/Mask_(computing))
- [JavaScript ビット演算子](https://www.w3schools.com/js/js_bitwise.asp)