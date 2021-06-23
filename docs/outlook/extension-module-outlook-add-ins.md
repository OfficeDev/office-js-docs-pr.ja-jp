---
title: Outlook のモジュール拡張機能アドイン
description: Outlook の内部で実行することで、ユーザーが Outlook から簡単にビジネスの情報や生産性ツールにアクセスできるようにするアプリケーションを作成します。
ms.date: 05/27/2020
localization_priority: Normal
ms.openlocfilehash: 3a02e93375f1c0872790d050382a14bc2c324cef
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/23/2021
ms.locfileid: "53077086"
---
# <a name="module-extension-outlook-add-ins"></a>Outlook のモジュール拡張機能アドイン

モジュール拡張機能アドインは、Outlook のナビゲーション バーのメール、タスク、および予定表の横に表示されます。 モジュール拡張機能は、メールと予定の情報のみ使用することに限定されていません。 Outlook の内部で実行することで、ユーザーが Outlook から簡単にビジネスの情報や生産性ツールにアクセスできるようにするアプリケーションを作成できます。

> [!NOTE]
> モジュール拡張機能は、Windows 用 Outlook 2016 以降でのみサポートされています。  

## <a name="open-a-module-extension"></a>モジュール拡張機能を開く

モジュール拡張機能を開くには、ユーザーは Outlook ナビゲーション バーのモジュール名またはアイコンをクリックします。ユーザーがコンパクト ナビゲーションを選択している場合、ナビゲーション バーには拡張機能がロードされていることを示すアイコンが表示されます。

![Outlook にモジュール拡張機能が読み込まれているときのコンパクト ナビゲーション バーを示します。](../images/outlook-module-navigationbar-compact.png)

ユーザーがコンパクト ナビゲーションを使用していない場合、ナビゲーション バーは 2 通りの見え方をします。 1 つの拡張機能が読み込まれている場合、そのアドインの名前が表示されます。

![Outlook にモジュール拡張機能が 1 つ読み込まれているときの拡張ナビゲーション バーを示します。](../images/outlook-module-navigationbar-one.png)

複数のアドインが読み込まれている場合は、**[アドイン]** という文字が表示されます。どちらをクリックしても、拡張機能のユーザー インターフェイスが開きます。

![Outlook にモジュール拡張機能が複数読み込まれている場合の拡張ナビゲーション バーを示します。](../images/outlook-module-navigationbar-more.png)

拡張機能をクリックすると、組み込みのモジュールは Outlook によってカスタム モジュールに置き換えられ、そのアドインはユーザーが対話的に操作できるようになります。 アドインでは Outlook JavaScript API のすべての機能を使用できます。また、Outlook のリボンにコマンド ボタンを作成することで、アドインの内容を対話的に操作できます。 以下のスクリーンショットは、Outlook のナビゲーション バーに統合され、アドインの内容を更新するリボン コマンドを有するアドインを示しています。

![モジュール拡張機能のユーザー インターフェイスを表示します。](../images/outlook-module-extension.png)

## <a name="example"></a>例

次に示すマニフェスト ファイルのセクションでは、モジュール拡張機能を定義しています。

```xml
<!-- Add Outlook module extension point -->
<VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides"
                  xsi:type="VersionOverridesV1_0">
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1"
                    xsi:type="VersionOverridesV1_1">

    <!-- Begin override of existing elements -->
    <Description resid="residVersionOverrideDesc" />

    <Requirements>
      <bt:Sets DefaultMinVersion="1.3">
        <bt:Set Name="Mailbox" />
      </bt:Sets>
    </Requirements>
    <!-- End override of existing elements -->

    <Hosts>
      <Host xsi:type="MailHost">
        <DesktopFormFactor>
          <!-- Set the URL of the file that contains the
                JavaScript function that controls the extension -->
          <FunctionFile resid="residFunctionFileUrl" />

          <!--New Extension Point - Module for a ModuleApp -->
          <ExtensionPoint xsi:type="Module">
            <SourceLocation resid="residExtensionPointUrl" />
            <Label resid="residExtensionPointLabel" />

            <CommandSurface>
              <CustomTab id="idTab">
                <Group id="idGroup">
                  <Label resid="residGroupLabel" />

                  <Control xsi:type="Button" id="group.changeToAssociate">
                    <Label resid="residChangeToAssociateLabel" />
                    <Supertip>
                      <Title resid="residChangeToAssociateLabel" />
                      <Description resid="residChangeToAssociateDesc" />
                    </Supertip>
                    <Icon>
                      <bt:Image size="16" resid="residAssociateIcon16" />
                      <bt:Image size="32" resid="residAssociateIcon32" />
                      <bt:Image size="80" resid="residAssociateIcon80" />
                    </Icon>
                    <Action xsi:type="ExecuteFunction">
                      <FunctionName>changeToAssociateRate</FunctionName>
                    </Action>
                  </Control>
                  
              </Group>
                <Label resid="residCustomTabLabel" />
              </CustomTab>
            </CommandSurface>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>
    </Hosts>

    <Resources>
      <bt:Images>
        <bt:Image id="residAddinIcon16" 
                  DefaultValue="https://localhost:8080/Executive-16.png" />
        <bt:Image id="residAddinIcon32" 
                  DefaultValue="https://localhost:8080/Executive-32.png" />
        <bt:Image id="residAddinIcon80" 
                  DefaultValue="https://localhost:8080/Executive-80.png" />
      
        <bt:Image id="residAssociateIcon16" 
                  DefaultValue="https://localhost:8080/Associate-16.png" />
        <bt:Image id="residAssociateIcon32" 
                  DefaultValue="https://localhost:8080/Associate-32.png" />
        <bt:Image id="residAssociateIcon80" 
                  DefaultValue="https://localhost:8080/Associate-80.png" />
      </bt:Images>

      <bt:Urls>
        <bt:Url id="residFunctionFileUrl" 
                DefaultValue="https://localhost:8080/" />
        <bt:Url id="residExtensionPointUrl" 
                DefaultValue="https://localhost:8080/" />
      </bt:Urls>

      <!--Short strings must be less than 30 characters long -->
      <bt:ShortStrings>
        <bt:String id="residExtensionPointLabel" 
                    DefaultValue="Billable Hours" />
        <bt:String id="residGroupLabel" 
                    DefaultValue="Change billing rate" />
        <bt:String id="residCustomTabLabel" 
                    DefaultValue="Billable hours" />

        <bt:String id="residChangeToAssociateLabel" 
                    DefaultValue="Associate" />
      </bt:ShortStrings>

      <bt:LongStrings>
        <bt:String id="residVersionOverrideDesc" 
                    DefaultValue="Version override description" />

        <bt:String id="residChangeToAssociateDesc" 
                    DefaultValue="Change to the associate billing rate: $127/hr" />
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
</VersionOverrides>
```

## <a name="see-also"></a>関連項目

- [Outlook アドインのマニフェスト](manifests.md)
- [Outlook のアドイン コマンド](add-in-commands-for-outlook.md)
- [Outlook モジュール拡張機能 "請求対象時間" のサンプル](https://github.com/OfficeDev/Outlook-Add-in-JavaScript-ModuleExtension)
