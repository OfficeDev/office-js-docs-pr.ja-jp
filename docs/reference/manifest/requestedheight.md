# <a name="requestedheight-element"></a><span data-ttu-id="52fb5-101">RequestedHeight 要素</span><span class="sxs-lookup"><span data-stu-id="52fb5-101">RequestedHeight element</span></span>

<span data-ttu-id="52fb5-102">コンテンツのアドインまたはメール アドインの初期の高さ (ピクセル単位で) を指定します。</span><span class="sxs-lookup"><span data-stu-id="52fb5-102">Specifies the initial height (in pixels) of a content add-in or mail add-in.</span></span> 

<span data-ttu-id="52fb5-103">**アドインの種類:** コンテンツ</span><span class="sxs-lookup"><span data-stu-id="52fb5-103">**Add-in type:** Content, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="52fb5-104">構文</span><span class="sxs-lookup"><span data-stu-id="52fb5-104">Syntax</span></span>

```XML
<RequestedHeight>integer</RequestedHeight>
```

## <a name="contained-in"></a><span data-ttu-id="52fb5-105">次に含まれる:</span><span class="sxs-lookup"><span data-stu-id="52fb5-105">Contained in:</span></span>

- <span data-ttu-id="52fb5-106">[DefaultSettings](defaultsettings.md) (アドインのコンテンツ) の値は、32 から 1000 の間にすることができます</span><span class="sxs-lookup"><span data-stu-id="52fb5-106">[DefaultSettings](defaultsettings.md) (Content add-ins) with a value that can be between 32 and 1000</span></span>
- <span data-ttu-id="52fb5-107">[DesktopSettings](desktopsettings.md)および[TabletSettings](tabletsettings.md) (メール アドインの場合) の値は、32 ~ 450 を指定できます。</span><span class="sxs-lookup"><span data-stu-id="52fb5-107">[DesktopSettings](desktopsettings.md) and [TabletSettings](tabletsettings.md) (Mail add-ins) with a value that can be between 32 and 450</span></span>
- <span data-ttu-id="52fb5-108">[ExtensionPoint](extensionpoint.md) (コンテキスト メール アドインの場合) の値は 140 と **DetectedEntity** の拡張点の 450 の間で、32 と **CustomPane** の拡張点の 450 の間で、</span><span class="sxs-lookup"><span data-stu-id="52fb5-108">[ExtensionPoint](extensionpoint.md) (Contextual mail add-ins) with a value that can be between 140 and 450 for the **DetectedEntity** extension point and between 32 and 450 for the **CustomPane** extension point</span></span>