---
title: Outlook アドインで代理人アクセスのシナリオを有効にする
description: 代理人アクセスについて簡単に説明し、アドインサポートを構成する方法について説明します。
ms.date: 01/14/2020
localization_priority: Normal
ms.openlocfilehash: 6cee68af9efc02bbb474effaba1a898511aea531
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/20/2020
ms.locfileid: "42166584"
---
# <a name="enable-delegate-access-scenarios-in-an-outlook-add-in"></a>Outlook アドインで代理人アクセスのシナリオを有効にする

メールボックスの所有者は代理人アクセス機能を使用して、[他のユーザーが自分のメールと予定表を管理できるよう](https://support.office.com/article/allow-someone-else-to-manage-your-mail-and-calendar-41c40c04-3bd1-4d22-963a-28eafec25926)にすることができます。 この記事では、Office JavaScript API でサポートされている代理人アクセス許可を指定し、Outlook アドインで代理人アクセスのシナリオを有効にする方法について説明します。

> [!IMPORTANT]
> 現在、Outlook on Mac、Android、iOS では、代理人アクセスは利用できません。 この機能は、今後利用可能になる可能性があります。
>
> この機能のサポートは、要件セット1.8 で導入されました。 この要件セットをサポートする [クライアントおよびプラットフォーム](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) を参照してください。

## <a name="supported-permissions-for-delegate-access"></a>代理人アクセスに対してサポートされるアクセス許可

次の表では、Office JavaScript API でサポートされている代理人アクセス許可について説明します。

|アクセス許可|値|説明|
|---|---:|---|
|読み取り|1 (000001)|アイテムを読み取ることができます。|
|書き込み|2 (000010)|アイテムを作成できます。|
|DeleteOwn|4 (000100)|は、自分で作成したアイテムのみを削除できます。|
|DeleteAll|8 (001000)|任意のアイテムを削除できます。|
|EditOwn|16 (010000)|は、自分で作成したアイテムのみを編集できます。|
|EditAll|32 (10万)|任意のアイテムを編集できます。|

> [!NOTE]
> 現在、API は既存の代理人アクセス許可の取得をサポートしていますが、代理人アクセス許可は設定しません。

[DelegatePermissions](/javascript/api/outlook/office.mailboxenums.delegatepermissions)オブジェクトは、デリゲートのアクセス許可を示すために、ビットマスクを使用して実装されます。 ビットマスク内の各位置は特定のアクセス許可を表し、それ`1`が設定されている場合は代理人にそれぞれのアクセス許可が付与されます。 たとえば、右側の2番目のビットが`1`の場合、デリゲートには**書き込み**アクセス許可があります。 この記事で後述する「[代理人として操作を実行](#perform-an-operation-as-delegate)する」の特定のアクセス許可を確認する方法の例を確認できます。

## <a name="sync-across-mailbox-clients"></a>メールボックスクライアント間での同期

通常、所有者のメールボックスに対する代理人の更新は、メールボックス間で即時に同期されます。

ただし、アドインが REST または EWS 操作を使用してアイテムの拡張プロパティを設定する場合、そのような変更は同期に数時間かかることがあります。このような遅延を避けるには、 [CustomProperties](/javascript/api/outlook/office.customproperties)オブジェクトと関連する api を使用することをお勧めします。 詳細については、記事「Outlook アドインでメタデータを取得および設定する」の「[カスタムプロパティ」セクション](metadata-for-an-outlook-add-in.md#custom-data-per-item-in-a-mailbox-custom-properties)を参照してください。

## <a name="configure-the-manifest"></a>マニフェストを構成する

アドインで代理人アクセスのシナリオを有効にするには、親要素`DesktopFormFactor`のマニフェスト内の`true` [supportssharedfolders](../reference/manifest/supportssharedfolders.md)要素をに設定する必要があります。 現在、他のフォームファクターはサポートされていません。

次の例は、 `SupportsSharedFolders`マニフェストのセクション`true`内に設定された要素を示しています。

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

アイテムの共有プロパティは、新規作成または閲覧モードで取得できます。そのためには、 [getSharedPropertiesAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)メソッドを呼び出します。 これにより、現在、代理人のアクセス許可、所有者の電子メールアドレス、REST API のベース URL、ターゲットメールボックスを提供する[Sharedproperties](/javascript/api/outlook/office.sharedproperties)オブジェクトが返されます。

次の例は、メッセージまたは予定の共有プロパティを取得する方法、代理人が**書き込み**アクセス許可を持っているかどうかを確認する方法、および REST 呼び出しを行う方法を示しています。

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

## <a name="see-also"></a>関連項目

- [自分のメールと予定表の管理を他のユーザーに許可する](https://support.office.com/article/allow-someone-else-to-manage-your-mail-and-calendar-41c40c04-3bd1-4d22-963a-28eafec25926)
- [Office365 での予定表の共有](https://support.office.com/article/calendar-sharing-in-office-365-b576ecc3-0945-4d75-85f1-5efafb8a37b4)
- [マニフェスト要素の注文方法](../develop/manifest-element-ordering.md)
- [マスク (コンピューティング)](https://en.wikipedia.org/wiki/Mask_(computing))
- [JavaScript ビット演算子](https://www.w3schools.com/js/js_bitwise.asp)