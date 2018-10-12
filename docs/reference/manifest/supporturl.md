# <a name="supporturl-element"></a><span data-ttu-id="5712f-101">SupportUrl 要素</span><span class="sxs-lookup"><span data-stu-id="5712f-101">SupportUrl element</span></span>

<span data-ttu-id="5712f-102">アドインのサポート情報を提供するページの URL を指定します。</span><span class="sxs-lookup"><span data-stu-id="5712f-102">Specifies the URL of a page that provides support information for your add-in.</span></span>

## <a name="syntax"></a><span data-ttu-id="5712f-103">構文</span><span class="sxs-lookup"><span data-stu-id="5712f-103">Syntax</span></span>

```XML
<OfficeApp>
...
  <IconUrl DefaultValue="https://contoso.com/assets/icon-32.png" />
  <HighResolutionIconUrl DefaultValue="https://contoso.com/assets/hi-res-icon.png"/>
  
  
  <SupportUrl DefaultValue="https://contoso.com/support " />
  
  
  <AppDomains>
  ...
  </AppDomains>
...
</OfficeApp>
```

## <a name="contained-in"></a><span data-ttu-id="5712f-104">次に含まれる:</span><span class="sxs-lookup"><span data-stu-id="5712f-104">Contained in:</span></span>

[<span data-ttu-id="5712f-105">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="5712f-105">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="5712f-106">含めることができるもの:</span><span class="sxs-lookup"><span data-stu-id="5712f-106">Can contain:</span></span>

|  <span data-ttu-id="5712f-107">要素</span><span class="sxs-lookup"><span data-stu-id="5712f-107">Element</span></span> | <span data-ttu-id="5712f-108">必須</span><span class="sxs-lookup"><span data-stu-id="5712f-108">Required</span></span> | <span data-ttu-id="5712f-109">説明</span><span class="sxs-lookup"><span data-stu-id="5712f-109">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="5712f-110">オーバーライド</span><span class="sxs-lookup"><span data-stu-id="5712f-110">Override</span></span>](override.md)   | <span data-ttu-id="5712f-111">いいえ</span><span class="sxs-lookup"><span data-stu-id="5712f-111">No</span></span> | <span data-ttu-id="5712f-112">追加のロケール URL の設定を指定します。</span><span class="sxs-lookup"><span data-stu-id="5712f-112">Specifies the setting for additional locale urls</span></span> |

## <a name="attributes"></a><span data-ttu-id="5712f-113">属性</span><span class="sxs-lookup"><span data-stu-id="5712f-113">Attributes</span></span>

|<span data-ttu-id="5712f-114">**属性**</span><span class="sxs-lookup"><span data-stu-id="5712f-114">**Attribute**</span></span>|<span data-ttu-id="5712f-115">**型**</span><span class="sxs-lookup"><span data-stu-id="5712f-115">**Type**</span></span>|<span data-ttu-id="5712f-116">**必須**</span><span class="sxs-lookup"><span data-stu-id="5712f-116">**Required**</span></span>|<span data-ttu-id="5712f-117">**説明**</span><span class="sxs-lookup"><span data-stu-id="5712f-117">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="5712f-118">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="5712f-118">DefaultValue</span></span>|<span data-ttu-id="5712f-119">URL</span><span class="sxs-lookup"><span data-stu-id="5712f-119">URL</span></span>|<span data-ttu-id="5712f-120">必須</span><span class="sxs-lookup"><span data-stu-id="5712f-120">required</span></span>|<span data-ttu-id="5712f-121">この設定の既定値を指定します。この値は、[DefaultLocale](defaultlocale.md) 要素に指定されるロケールを対象としています。</span><span class="sxs-lookup"><span data-stu-id="5712f-121">Specifies the default value for this setting, expressed for the locale specified in the [DefaultLocale](defaultlocale.md) element.</span></span>|
