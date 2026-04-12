---
title: "UltimateHtmlViewer Overview"
description: "How UltimateHtmlViewer fits SharePoint-hosted HTML report experiences."
layout: docs
---

UltimateHtmlViewer is useful when generated HTML reports or static app bundles need to be hosted inside SharePoint without breaking relative navigation or forcing users into raw file URLs.

## Common fit

- publish generated HTML reports into SharePoint
- keep report navigation inside one modern page
- support deep links via `?uhvPage=...`
- apply URL policy controls such as `StrictTenant` or `Allowlist`
- use a reusable SPFx web part instead of creating a custom host for every report bundle

## Good operating pattern

Use a local profile file under `ignore/` for tenant values and client IDs. Keep deployment details out of committed examples unless they are placeholders.

