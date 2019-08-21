---
title: マニフェストファイルの Authorization 要素
description: ''
ms.date: 08/12/2019
localization_priority: Normal
ms.openlocfilehash: cc3b80e0e02eca9c197b82931a6f2011ba385d57
ms.sourcegitcommit: da8e6148f4bd9884ab9702db3033273a383d15f0
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/20/2019
ms.locfileid: "36477944"
---
# <a name="authorization-element"></a><span data-ttu-id="c9bab-102">Authorization 要素</span><span class="sxs-lookup"><span data-stu-id="c9bab-102">Authorization element</span></span>

<span data-ttu-id="c9bab-103">アドインの web アプリケーションが承認を必要とする外部リソースと、必要なアクセス許可を指定します。</span><span class="sxs-lookup"><span data-stu-id="c9bab-103">Specifies the external resources that the add-in's web application needs authorization to and the required permissions.</span></span>

<span data-ttu-id="c9bab-104">**承認**は、マニフェスト内の[承認](authorizations.md)要素の子要素です。</span><span class="sxs-lookup"><span data-stu-id="c9bab-104">**Authorization** is a child element of the [Authorizations](authorizations.md) element in the manifest.</span></span>

## <a name="child-elements"></a><span data-ttu-id="c9bab-105">子要素</span><span class="sxs-lookup"><span data-stu-id="c9bab-105">Child elements</span></span>

|  <span data-ttu-id="c9bab-106">要素</span><span class="sxs-lookup"><span data-stu-id="c9bab-106">Element</span></span> |  <span data-ttu-id="c9bab-107">必須</span><span class="sxs-lookup"><span data-stu-id="c9bab-107">Required</span></span>  |  <span data-ttu-id="c9bab-108">説明</span><span class="sxs-lookup"><span data-stu-id="c9bab-108">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="c9bab-109">**Resource**</span><span class="sxs-lookup"><span data-stu-id="c9bab-109">**Resource**</span></span>  |  <span data-ttu-id="c9bab-110">はい</span><span class="sxs-lookup"><span data-stu-id="c9bab-110">Yes</span></span>   |  <span data-ttu-id="c9bab-111">外部リソースの URL を指定します。</span><span class="sxs-lookup"><span data-stu-id="c9bab-111">Specifies the URL of the external resource.</span></span>|
|  [<span data-ttu-id="c9bab-112">Scope</span><span class="sxs-lookup"><span data-stu-id="c9bab-112">Scopes</span></span>](scopes.md)                |  <span data-ttu-id="c9bab-113">はい</span><span class="sxs-lookup"><span data-stu-id="c9bab-113">Yes</span></span>  |  <span data-ttu-id="c9bab-114">アドインがリソースに対して必要とするアクセス許可を指定します。</span><span class="sxs-lookup"><span data-stu-id="c9bab-114">Specifies the permissions that the add-in needs to the resource.</span></span>  |

## <a name="example"></a><span data-ttu-id="c9bab-115">例</span><span class="sxs-lookup"><span data-stu-id="c9bab-115">Example</span></span>

```xml
<OfficeApp>
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
    ...
    <WebApplicationInfo>
      <Id>12345678-abcd-1234-efab-123456789abc</Id>
      <Resource>api://myDomain.com/12345678-abcd-1234-efab-123456789abc</Resource>
      <Scopes>
        <Scope>Files.Read.All</Scope>
        <Scope>offline_access</Scope>
        <Scope>openid</Scope>
        <Scope>profile</Scope>
      </Scopes>
      <Authorizations>
        <Authorization>
          <Resource>https://api.contoso.com</Resource>
            <Scopes>
              <Scope>profile</Scope>
          </Scopes>
        </Authorization>
      </Authorizations>
    </WebApplicationInfo>
  </VersionOverrides>
...
</OfficeApp>
```
