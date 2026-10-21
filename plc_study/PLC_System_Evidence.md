# PLC Study: a real telecom system, with real evidence

Prepared for the Telecom Egypt PLC study team. Research date: October 2026.

> **الخلاصة بالعربي:** النظام المقترح هو **Netcracker Product Management** (وهو عبارة عن Product Catalog + Product Lifecycle Management). ده نظام حقيقي شغال فعلًا عند شركات اتصالات معروفة، وعليه أدلة منشورة: Zain السعودية قللت وقت تطوير المنتج 50%، وT-Mobile هولندا اختارته مخصوص علشان الـ Product Catalog والـ PLC، وكمان e& الإمارات وdu. وفيه مصدر خارجي بيقول إن المصرية للاتصالات نفسها اشترت Netcracker BSS/OSS سنة 2013، ودي معلومة محتاجة نتأكد منها من فريق الـ IT عندنا. كل معلومة تحت ليها مصدر، ومكتوب جنبها نوع الدليل.

---

## 1. Recommended system

| | |
|---|---|
| **System** | Netcracker Product Management: **Product Catalog** + **Product Lifecycle Management (PLM)** |
| **Vendor** | Netcracker Technology (USA), a wholly owned subsidiary of NEC Corporation (Japan) |
| **Part of** | Netcracker Digital BSS, now sold as part of *Digital Commerce & Monetization* |
| **Why this one** | (1) Its product is literally named "Product Lifecycle Management", which matches our PLC scope. (2) It has named telecom deployments with published results. (3) Its strongest references are in our region (Saudi Arabia, UAE). (4) A third-party source links Netcracker to Telecom Egypt itself (see section 3). |

### Why I dropped Oracle Launch Cloud Service (the first choice)

Oracle Launch has the best public *documentation* of the three systems I checked. However, I searched for a named operator that publicly uses it and found none. Oracle's telecom press releases are about its billing and charging products (for example SmarTone, Hong Kong, 2024), not Launch. Without a real customer, it fails the "real evidence" requirement. It is still useful as a reference for *how* a PLC system is designed internally, because its manuals are public.

---

## 2. Evidence: operators actually using it

Evidence types, from strongest to weakest:
- **Operator statement:** a named operator executive is quoted in the source.
- **Vendor case study:** published by the vendor, with measured results.
- **Vendor press release:** an announcement, usually at contract signature.
- **Third-party database:** not confirmed by the vendor or the operator.

| Operator | Country | What was deployed | Published result | Evidence type | Source |
|---|---|---|---|---|---|
| **Zain KSA** | Saudi Arabia | Full BSS/OSS, including **Product Catalog**, CPQ, convergent charging, partner management | **50% decrease in product development time** and in customer line activation time; whole program done in under 3 years | Vendor case study with operator quotes (EVP IT, CEO, CTO) | [Netcracker Zain KSA case study (PDF)](https://netcracker.com/documents/pdf/netcracker_zain_case-study_260918.pdf); selection announced 4 Oct 2021: [press release](https://www.netcracker.com/news/press-releases/netcracker-delivers-bss-transformation-for-zain-saudi-arabia) |
| **T-Mobile Netherlands** (Deutsche Telekom) | Netherlands | **Product Catalog + Product Lifecycle Management** specifically, replacing built-to-order systems | Goals: faster time-to-market and **fewer incorrect offers**; manage products "from development to production to retirement" | Vendor press release with operator quote (Manager Solution Design) | [Netcracker press release](https://www.netcracker.com/news/press-releases/t-mobile-netherlands-optimizes-end-to-end-product-management-with-netcracker-bss/) |
| **e& UAE** (Etisalat) | UAE | Full-stack BSS including **product catalog management**, CPQ, order management, CRM | Operator: "improving time to market and laying the groundwork for **catalog-driven offers**"; end-to-end digital onboarding for B2B | Vendor press release with operator quote (Acting CTIO) | [Partnership, 26 Oct 2023](https://netcracker.com/news/press-releases/etisalat-by-e-and-partners-with-netcracker-on-the-largest-full-stack-bss-transformation-project-in-the-middle-east); [milestone](https://www.netcracker.com/news/press-releases/netcracker-and-e-achieve-significant-milestone-in-bss-transformation-program); [case study (PDF)](https://netcracker.com/documents/pdf/netcracker_etisalat_case-study.pdf) |
| **du** | UAE | Digital BSS including the **Product Management** domain | Stated aim: improve time to market, migrate from legacy | Vendor press release (contract stage, 12 Apr 2021) | [Netcracker press release](https://www.netcracker.com/news/press-releases/netcracker-to-deliver-bss-transformation-for-leading-middle-eastern-operator-du) |
| **Telecom Egypt** | Egypt | "NetCracker BSS/OSS", bought 2013, live 2014 | None published | **Third-party database only; not confirmed** | [AppsRunTheWorld: Telecom Egypt](https://www.appsruntheworld.com/customers-database/customers/view/telecom-egypt-egypt) |

**Important for us:** if Telecom Egypt really runs Netcracker, our own BSS/IT team may already have access to it, or to the vendor. That would let us see the real screens and workflow instead of relying on marketing material. This is the first thing to ask internally.

Other Telecom Egypt context found (not about product catalogs): Huawei states that its BSS underpins the WE-branded bundles ([Huawei WinWin](https://www.huawei.com/en/huaweitech/publication/winwin/34/telecom-egypt-powering-diversification-with-ict)). Ericsson Mediation was selected in 2025 for customer communications ([Ericsson](https://www.ericsson.com/en/press-releases/5/2025/telecom-egypt-modernizes-customer-experiences-with-ericsson-mediation)).

---

## 3. What the system does (mapped to the features in our assignment)

Source for the module descriptions: Netcracker's Product Management product page (retrieved through search in Oct 2026; the page has since moved) and its [Commerce Management page](https://netcracker.com/portfolio/products/digital-commerce-monetization/commerce-management).

**Two components:**
1. **Product Catalog:** "a unified holistic view and easy access to up-to-date product information (pricing, discounts, fees, loyalty programs, taxes), supporting sales, ordering and billing processes via industry-standard integration."
2. **Product Lifecycle Management:** "end-to-end lifecycle management from **marketing concept development** and product configuration to **product testing, test launch** and subsequent **mass production**. It also fully supports **product versioning, upgrades and retirement**."

**Supporting capabilities stated by Netcracker:** pre-defined product templates; federation of third-party partner catalogs into one interface; analysis of product consumption to target offers; consistency of product data across all channels; GenAI-assisted configuration; business users (not IT) configure telco and non-telco offers.

| Feature from our brief | What Netcracker says it does | Evidence level |
|---|---|---|
| Product Ideation | PLM starts at "marketing concept development" | Vendor product page |
| Product Definition | Template-based configuration from reusable product components | Vendor product page |
| Approval Workflow | Configurable approval workflows (documented for CPQ quotes); PLM governance stages | Vendor whitepaper (CPQ); **product-approval detail not public** |
| Product Catalog | Unified catalog: prices, discounts, fees, loyalty, taxes; one view for all channels | Vendor product page + Zain, T-Mobile NL deployments |
| Pricing Management | Pricing by market, customer category and channel; discounts and incentives | Vendor whitepaper (CPQ) |
| Product Versioning | "Fully supports product versioning" | Vendor product page |
| Role-Based Access | Not described publicly | **Need vendor demo** |
| Product Launch | Testing, then test launch, then mass production | Vendor product page |
| CRM Integration | Catalog feeds sales, ordering and billing via industry-standard (TM Forum) integration; CRM is in the same suite | Vendor pages + e& deployment (CRM + catalog) |
| Product Retirement | "Supports ... retirement"; T-Mobile NL: "from development to production to retirement" | Vendor page + operator press release |
| Product Migration | "Upgrades" supported; CPQ handles "change and up-sell scenarios" | Vendor pages; **customer-migration detail not public** |

**Honest gap:** vendors do not publish detailed screens for approvals, roles or migration. These can only be confirmed through a vendor demo, or through our own IT team if Telecom Egypt already has the product.

---

## 4. Real use cases (published, not invented)

**Use case 1: Zain KSA, faster product development (Saudi Arabia).**
- **Problem:** a legacy IT stack that slowed 5G monetization and expansion into fintech, IoT and cloud.
- **What they did:** replaced the full stack, including Product Catalog and CPQ, then migrated customers to the new cloud-based systems.
- **Result:** product development time was cut by 50%, and the program finished in under 3 years.
- **Relevance to us:** a regional operator with a similar scale and market.

**Use case 2: T-Mobile Netherlands, a dedicated PLC project.**
- **Problem:** out-of-date, built-to-order (custom) systems for managing products.
- **What they did:** bought Netcracker *only* for Product Catalog + Product Lifecycle Management, as part of a wider IT program.
- **Goal:** one place to manage products "from development to production to retirement", faster time to market, and fewer incorrect offers.
- **Relevance to us:** this is the closest real example of exactly what our team is studying, a PLC system as its own project.

**Use case 3: e& UAE, catalog-driven offers for business customers.**
- **Problem:** product catalog, quoting and ordering needed modernizing for B2B and over-the-top services.
- **What they did:** product catalog optimization, CPQ, order capture and fulfillment, and a centralized platform for over-the-top services.
- **Result (operator quote):** improved time to market, end-to-end digital onboarding of B2B customers, and "groundwork for catalog-driven offers".

---

## 5. Two alternatives with strong evidence (for the other two team members)

The brief says each of the three team members studies a different system. These two also have real, checkable evidence:

**Amdocs CatalogONE**
- **Sprint, USA (now T-Mobile US):** an *independent* TM Forum awards case study. Sprint replaced 22 channel catalogs with a single catalog and expected a 50% cut in time to market. Business users define offers themselves, so the people needed per offer dropped from 20–30 to 3–4. [TM Forum case study (PDF)](https://www.tmforum.org/wp-content/uploads/2021/02/TM-Forum-Excellence-Awards-Case-Study-Cloud_Sprint.pdf)
- **KT, South Korea:** migrated its full product catalog (all offers and price plans) to CatalogONE in 2020. The press release cites "approvals and notification management" for business users. [GlobeNewswire, Nov 2020](https://www.globenewswire.com/news-release/2020/11/10/2124219/0/en/KT-Successfully-Implements-Amdocs-CatalogONE-Cloud-Native-Solution-to-Accelerate-Introduction-of-New-5G-Offerings.html)
- **An unnamed European quad-play operator:** new offers went from up to 9 months to 1–2 weeks. This is reported by an independent analyst, Appledore Research. [Appledore solution profile (PDF)](https://www.amdocs.com/sites/default/files/2021-07/Amdocs-CatalogONE-Appledore-Solution-Profile.pdf)
- **Also:** Globe Telecom (Philippines, 2020). Amdocs' own 2025 material describes a PLM module "from ideation through creation, implementation and up to end-of-life", with workflow and approval management.

**Oracle Communications Launch Cloud Service**
- **Strength:** fully public manuals covering lifecycle statuses, approval workflow (sequential or parallel), 5 user roles, versioning, retirement, upgrade/downgrade migration rules, and publishing to CRM (Siebel). It is the best source for understanding the internal design. [User guide](https://docs.oracle.com/en/industries/communications/launch-cloud-service/26c/user-guide/)
- **Weakness:** no named telecom customer found publicly.

---

## 6. Next steps to strengthen the evidence

1. **Ask Telecom Egypt's BSS/IT team** whether the 2013 Netcracker BSS/OSS is still in use, and whether it includes the Product Catalog / PLM modules. If it does, ask for a walkthrough of the real system.
2. **Request a vendor demo** from Netcracker (Middle East office) for the gaps above: the approval screens, roles and permissions, and customer migration.
3. **Use TM Forum TMF620**, the Product Catalog Management API standard, as the neutral reference for the lifecycle statuses: in study, in design, in test, active, launched, retired, obsolete. All three vendors build on TM Forum standards. [TMF620](https://www.tmforum.org/open-digital-architecture/open-apis/product-catalog-management-api-TMF620/v5.0)

---

## Sources

- Netcracker: [Zain KSA case study](https://netcracker.com/documents/pdf/netcracker_zain_case-study_260918.pdf), [Zain KSA press release (2021)](https://www.netcracker.com/news/press-releases/netcracker-delivers-bss-transformation-for-zain-saudi-arabia), [T-Mobile Netherlands press release](https://www.netcracker.com/news/press-releases/t-mobile-netherlands-optimizes-end-to-end-product-management-with-netcracker-bss/), [e& partnership (2023)](https://netcracker.com/news/press-releases/etisalat-by-e-and-partners-with-netcracker-on-the-largest-full-stack-bss-transformation-project-in-the-middle-east), [e& milestone](https://www.netcracker.com/news/press-releases/netcracker-and-e-achieve-significant-milestone-in-bss-transformation-program), [du (2021)](https://www.netcracker.com/news/press-releases/netcracker-to-deliver-bss-transformation-for-leading-middle-eastern-operator-du), [Commerce Management](https://netcracker.com/portfolio/products/digital-commerce-monetization/commerce-management), [CPQ whitepaper](https://netcracker.com/assets/portfolio/solutions/intelligent-sales-automation/netcracker_cpq-whitepaper.pdf)
- Telecom Egypt context: [AppsRunTheWorld](https://www.appsruntheworld.com/customers-database/customers/view/telecom-egypt-egypt), [Huawei WinWin](https://www.huawei.com/en/huaweitech/publication/winwin/34/telecom-egypt-powering-diversification-with-ict), [Ericsson 2025](https://www.ericsson.com/en/press-releases/5/2025/telecom-egypt-modernizes-customer-experiences-with-ericsson-mediation)
- Amdocs: [TM Forum Sprint case study](https://www.tmforum.org/wp-content/uploads/2021/02/TM-Forum-Excellence-Awards-Case-Study-Cloud_Sprint.pdf), [Appledore profile](https://www.amdocs.com/sites/default/files/2021-07/Amdocs-CatalogONE-Appledore-Solution-Profile.pdf), [KT 2020](https://www.globenewswire.com/news-release/2020/11/10/2124219/0/en/KT-Successfully-Implements-Amdocs-CatalogONE-Cloud-Native-Solution-to-Accelerate-Introduction-of-New-5G-Offerings.html), [Globe 2020](https://www.globenewswire.com/news-release/2020/05/07/2029864/34518/en/globe-selects-amdocs-cloud-native-solution-to-transform-its-enterprise-operations.html), [Amdocs catalog e-book 2025](https://solutions.amdocs.com/rs/647-OJR-802/images/Amdocs-signaturecx-campaign-ebook-catalog-02-2025.pdf)
- Oracle: [Launch user guide](https://docs.oracle.com/en/industries/communications/launch-cloud-service/26c/user-guide/), [SmarTone 2024 (billing, not Launch)](https://www.prnewswire.com/news-releases/smartone-selects-oracle-to-transform-and-adopt-cloud-based-charging-and-billing-system-at-scale-to-accelerate-the-delivery-of-new-services-302050145.html)
