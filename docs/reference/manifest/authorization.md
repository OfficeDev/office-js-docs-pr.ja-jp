---
title: マニフェストファイルの Authorization 要素
description: アドインの web アプリケーションが承認を必要とする外部リソースと、必要なアクセス許可を指定します。
ms.date: 08/12/2019
localization_priority: Normal
ms.openlocfilehash: cece0934eb9db3175b173e97d7ab478827b7cda2
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718441"
---
# <a name="authorization-element"></a><span data-ttu-id="71195-103">Authorization 要素</span><span class="sxs-lookup"><span data-stu-id="71195-103">Authorization element</span></span>

<span data-ttu-id="71195-104">アドインの web アプリケーションが承認を必要とする外部リソースと、必要なアクセス許可を指定します。</span><span class="sxs-lookup"><span data-stu-id="71195-104">Specifies the external resources that the add-in's web application needs authorization to and the required permissions.</span></span>

<span data-ttu-id="71195-105">**承認**は、マニフェスト内の[承認](authorizations.md)要素の子要素です。</span><span class="sxs-lookup"><span data-stu-id="71195-105">**Authorization** is a child element of the [Authorizations](authorizations.md) element in the manifest.</span></span>

## <a name="child-elements"></a><span data-ttu-id="71195-106">子要素</span><span class="sxs-lookup"><span data-stu-id="71195-106">Child elements</span></span>

|  <span data-ttu-id="71195-107">要素</span><span class="sxs-lookup"><span data-stu-id="71195-107">Element</span></span> |  <span data-ttu-id="71195-108">必須</span><span class="sxs-lookup"><span data-stu-id="71195-108">Required</span></span>  |  <span data-ttu-id="71195-109">説明</span><span class="sxs-lookup"><span data-stu-id="71195-109">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="71195-110">**Resource**</span><span class="sxs-lookup"><span data-stu-id="71195-110">**Resource**</span></span>  |  <span data-ttu-id="71195-111">はい</span><span class="sxs-lookup"><span data-stu-id="71195-111">Yes</span></span>   |  <span data-ttu-id="71195-112">外部リソースの URL を指定します。</span><span class="sxs-lookup"><span data-stu-id="71195-112">Specifies the URL of the external resource.</span></span>|
|  [<span data-ttu-id="71195-113">Scope</span><span class="sxs-lookup"><span data-stu-id="71195-113">Scopes</span></span>](scopes.md)                |  <span data-ttu-id="71195-114">はい</span><span class="sxs-lookup"><span data-stu-id="71195-114">Yes</span></span>  |  <span data-ttu-id="71195-115">アドインがリソースに対して必要とするアクセス許可を指定します。</span><span class="sxs-lookup"><span data-stu-id="71195-115">Specifies the permissions that the add-in needs to the resource.</span></span>  |

## <a name="example"></a><span data-ttu-id="71195-116">例</span><span class="sxs-lookup"><span data-stu-id="71195-116">Example</span></span>

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
