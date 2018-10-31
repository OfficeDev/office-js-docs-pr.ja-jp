# <a name="supportssharedfolders-element"></a>SupportsSharedFolders 要素

Outlook アドインが委任のシナリオで使用可能かどうかを定義します。  **SupportsSharedFolders** 要素は、[DesktopFormFactor](desktopformfactor.md) の子要素です。 既定で *false* に設定されています。

> [!IMPORTANT]
> この要素は、Exchange Online への [Outlook アドイン プレビュー要求セット](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md)でのみ使用可能です。 この要素を使用するアドインは、AppSource に発行または一元展開で展開できません。

**SupportsSharedFolders** 要素の例を次に示します。

```XML
<DesktopFormFactor>
  <FunctionFile resid="residDesktopFuncUrl" />
  <SupportsSharedFolders>true</SupportsSharedFolders>
  <ExtensionPoint xsi:type="PrimaryCommandSurface">
    <!-- information about this extension point -->
  </ExtensionPoint>

  <!-- You can define more than one ExtensionPoint element as needed -->

</DesktopFormFactor>
```
