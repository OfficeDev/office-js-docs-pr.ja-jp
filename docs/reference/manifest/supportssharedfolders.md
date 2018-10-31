# <a name="supportssharedfolders-element"></a><span data-ttu-id="e6c68-101">SupportsSharedFolders 要素</span><span class="sxs-lookup"><span data-stu-id="e6c68-101">SupportsSharedFolders element</span></span>

<span data-ttu-id="e6c68-102">Outlook アドインが委任のシナリオで使用可能かどうかを定義します。</span><span class="sxs-lookup"><span data-stu-id="e6c68-102">Defines whether the Outlook add-in is available in delegate scenarios.</span></span> <span data-ttu-id="e6c68-103"> *\*SupportsSharedFolders** 要素は、[DesktopFormFactor](desktopformfactor.md) の子要素です。</span><span class="sxs-lookup"><span data-stu-id="e6c68-103">The **ExtensionPoint** element is a child element of [AllFormFactors, DesktopFormFactor or MobileFormFactor](desktopformfactor.md).</span></span> <span data-ttu-id="e6c68-104">既定で *false* に設定されています。</span><span class="sxs-lookup"><span data-stu-id="e6c68-104">It is set to *false* by default.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="e6c68-105">この要素は、Exchange Online への [Outlook アドイン プレビュー要求セット](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md)でのみ使用可能です。</span><span class="sxs-lookup"><span data-stu-id="e6c68-105">This element is only available in the [Outlook add-ins Preview requirement set](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) against Exchange Online.</span></span> <span data-ttu-id="e6c68-106">この要素を使用するアドインは、AppSource に発行または一元展開で展開できません。</span><span class="sxs-lookup"><span data-stu-id="e6c68-106">Add-ins that use this element cannot be published to AppSource or deployed via centralized deployment.</span></span>

<span data-ttu-id="e6c68-107">**SupportsSharedFolders** 要素の例を次に示します。</span><span class="sxs-lookup"><span data-stu-id="e6c68-107">The following is an example of the **FunctionFile** element.</span></span>

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
