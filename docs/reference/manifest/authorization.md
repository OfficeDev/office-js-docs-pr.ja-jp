---
title: マニフェストファイルの Authorization 要素
description: アドインの web アプリケーションが承認を必要とする外部リソースと、必要なアクセス許可を指定します。
ms.date: 08/12/2019
localization_priority: Normal
ms.openlocfilehash: b8d3dd31a212a7de00ff4dbf263e8593a8ec2898
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/28/2020
ms.locfileid: "47294263"
---
# <a name="authorization-element"></a><span data-ttu-id="5176c-103">Authorization 要素</span><span class="sxs-lookup"><span data-stu-id="5176c-103">Authorization element</span></span>

<span data-ttu-id="5176c-104">アドインの web アプリケーションが承認を必要とする外部リソースと、必要なアクセス許可を指定します。</span><span class="sxs-lookup"><span data-stu-id="5176c-104">Specifies the external resources that the add-in's web application needs authorization to and the required permissions.</span></span>

<span data-ttu-id="5176c-105">**承認** は、マニフェスト内の [承認](authorizations.md) 要素の子要素です。</span><span class="sxs-lookup"><span data-stu-id="5176c-105">**Authorization** is a child element of the [Authorizations](authorizations.md) element in the manifest.</span></span>

## <a name="child-elements"></a><span data-ttu-id="5176c-106">子要素</span><span class="sxs-lookup"><span data-stu-id="5176c-106">Child elements</span></span>

|  <span data-ttu-id="5176c-107">要素</span><span class="sxs-lookup"><span data-stu-id="5176c-107">Element</span></span> |  <span data-ttu-id="5176c-108">必須</span><span class="sxs-lookup"><span data-stu-id="5176c-108">Required</span></span>  |  <span data-ttu-id="5176c-109">説明</span><span class="sxs-lookup"><span data-stu-id="5176c-109">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="5176c-110">**Resource**</span><span class="sxs-lookup"><span data-stu-id="5176c-110">**Resource**</span></span>  |  <span data-ttu-id="5176c-111">はい</span><span class="sxs-lookup"><span data-stu-id="5176c-111">Yes</span></span>   |  <span data-ttu-id="5176c-112">外部リソースの URL を指定します。</span><span class="sxs-lookup"><span data-stu-id="5176c-112">Specifies the URL of the external resource.</span></span>|
|  [<span data-ttu-id="5176c-113">Scope</span><span class="sxs-lookup"><span data-stu-id="5176c-113">Scopes</span></span>](scopes.md)                |  <span data-ttu-id="5176c-114">はい</span><span class="sxs-lookup"><span data-stu-id="5176c-114">Yes</span></span>  |  <span data-ttu-id="5176c-115">アドインがリソースに対して必要とするアクセス許可を指定します。</span><span class="sxs-lookup"><span data-stu-id="5176c-115">Specifies the permissions that the add-in needs to the resource.</span></span>  |

## <a name="example"></a><span data-ttu-id="5176c-116">例</span><span class="sxs-lookup"><span data-stu-id="5176c-116">Example</span></span>

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
