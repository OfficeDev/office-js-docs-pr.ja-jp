---
title: マニフェストファイルの承認要素
description: ''
ms.date: 08/12/2019
localization_priority: Normal
ms.openlocfilehash: 6a271423ddd549431c2f580e2793faab3c49090e
ms.sourcegitcommit: da8e6148f4bd9884ab9702db3033273a383d15f0
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/20/2019
ms.locfileid: "36477958"
---
# <a name="authorizations-element"></a><span data-ttu-id="83af5-102">承認要素</span><span class="sxs-lookup"><span data-stu-id="83af5-102">Authorizations element</span></span>

<span data-ttu-id="83af5-103">アドインの web アプリケーションが承認を必要とする外部リソースと、必要なアクセス許可を指定します。</span><span class="sxs-lookup"><span data-stu-id="83af5-103">Specifies the external resources that the add-in's web application needs authorization to and the required permissions.</span></span>

<span data-ttu-id="83af5-104">**承認**は、マニフェスト内の[webapplicationinfo](webapplicationinfo.md)要素の子要素です。</span><span class="sxs-lookup"><span data-stu-id="83af5-104">**Authorizations** is a child element of the [WebApplicationInfo](webapplicationinfo.md) element in the manifest.</span></span>

## <a name="child-elements"></a><span data-ttu-id="83af5-105">子要素</span><span class="sxs-lookup"><span data-stu-id="83af5-105">Child elements</span></span>

|  <span data-ttu-id="83af5-106">要素</span><span class="sxs-lookup"><span data-stu-id="83af5-106">Element</span></span> |  <span data-ttu-id="83af5-107">必須</span><span class="sxs-lookup"><span data-stu-id="83af5-107">Required</span></span>  |  <span data-ttu-id="83af5-108">説明</span><span class="sxs-lookup"><span data-stu-id="83af5-108">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="83af5-109">Authorization</span><span class="sxs-lookup"><span data-stu-id="83af5-109">Authorization</span></span>](authorization.md)                |  <span data-ttu-id="83af5-110">はい</span><span class="sxs-lookup"><span data-stu-id="83af5-110">Yes</span></span>     |   <span data-ttu-id="83af5-111">アドインの web アプリケーションが承認を必要とする外部リソースと、必要なスコープ (アクセス許可) を識別します。</span><span class="sxs-lookup"><span data-stu-id="83af5-111">Identifies an external resource that the add-in's web application needs authorization to, and the scopes (permissions) that it needs.</span></span> |

## <a name="example"></a><span data-ttu-id="83af5-112">例</span><span class="sxs-lookup"><span data-stu-id="83af5-112">Example</span></span>

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
