**English** | [한국어](research-methods.ko.md)

# Research methods for reducing information loss

This guide records the research behind the v0.4.0 reliability changes. The implementation choices are our engineering inferences, with tests on SlideFill scoped separately. [Feature algorithms](feature-algorithms.md) · [Verification limits](verification.md) · [Project experiments](reliability-history.md).

## Search and evidence

On 9 November 2026, Liner Scholar MCP was used for three broad search batches covering spreadsheet types/provenance, metamorphic testing/oracles, and contracts/numeric deviations, followed by focused title/topic searches. Only generic public research queries were sent. No source workbooks, user documents or library uploads were involved.

Liner returned usable body text for the first two papers below. Their excerpts were sufficient to examine the relevant mechanisms. Other rows rely on the returned abstracts. Two additional fetches reported that full text was unavailable despite a positive search availability flag; neither body is treated as read evidence. This is a focused engineering search, not an exhaustive systematic review.

| Primary source and Liner availability | What the evidence supports | Application inference and limitation |
|---|---|---|
| Abraham & Erwig, [Header and Unit Inference for Spreadsheets Through Spatial Analyses](https://doi.org/10.1109/VLHCC.2004.29), 2004. Full text retrieved; relevant excerpt inspected. | Headers determine structured units; incompatible inferred units can expose bad references. Header inference depends on spatial layout and combines several heuristics. | Carry the comparison's kind, period and exact value explicitly between modules. Equal period labels alone do not establish equivalent business measures. This change does not implement the paper's unit inference system. |
| Liu et al., [How Effectively Does Metamorphic Testing Alleviate the Oracle Problem?](https://doi.org/10.1109/TSE.2013.46), online 2013. Full text retrieved; relevant excerpt inspected. | Relations between several executions provide necessary properties when individual expected outputs are costly. The study checks relation validity and investigates diverse relations. | Test equivalent decimal spellings, positive scaling and preserved direction, together with fixed counterexamples. Preconditions matter: a relation can be invalid for a different operation or zero base. A passed relation is not a proof of all outputs. |
| Barr et al., [The Oracle Problem in Software Testing: A Survey](https://doi.org/10.1109/TSE.2014.2372785), 2015. Full text listed available; abstract inspected, body not fetched. | Specifications, contracts and metamorphic methods are sources of oracle information; informal domain knowledge can remain necessary. | Use independently stated rational expectations and canvas geometry checks alongside end-to-end saved-output checks. Agreement between two paths can still share an error. |
| Almasi et al., [Search-based detection of deviation failures in the migration of legacy spreadsheet applications](https://doi.org/10.1145/3213846.3213861), 2018. Full text listed unavailable; abstract inspected. | The study targets numerical deviations when spreadsheet formulas are migrated to Java, using two financial applications. | Target adjacent large integers, long decimals, tiny nonzero changes and extreme magnitudes. Its setting and fault-detection rates are not SlideFill accuracy estimates, and its tolerance policy is not adopted here. |
| Chen et al., [Metamorphic Testing: A Review of Challenges and Opportunities](https://doi.org/10.1145/3143561), 2018. Full text listed unavailable; abstract inspected. | Metamorphic testing combines input generation with result checks based on necessary relations. Identifying useful relations remains a research challenge. | Document why each transformation should preserve or predict a result, and combine relations with concrete false-accept and false-reject cases. Generated tests alone do not determine the intended business calculation. |

## Bounded application contract

The useful common thread is to preserve information before presentation, specify the boundary, and test it with an oracle suited to that boundary.

| Boundary | Required behavior | Evidence to collect |
|---|---|---|
| Fact number to optional model wording | Supported decimal tokens retain exact identity under bounded parsing. Equivalent spellings may agree; distinct values must not collapse through float display rounding. | Wrong adjacent integers and differing long decimals are rejected; supported grouping and zero-padding transformations remain accepted; excessive tokens are refused. |
| Chosen numeric prose to editor/save reuse | Apply the same narrow numeric-magnitude check against the current facts. Reject an invalid edit while preserving the previous choice; use a warning and calculated fallback for an invalid stored choice. | Editor acceptance/rejection and close behavior, saved-choice tampering, equivalent spellings and clearing the choice. |
| Calculated change to cross-file reading | Direction comes from an exact typed change, with explicit relative-change basis and period. Presentation text is not decoded back into the arithmetic input. | A tiny positive/negative change keeps its direction; positive scaling preserves relative change; rate point changes are not relabeled as relative changes. |
| Numeric series to contents sketch | Supported finite rational data produces bounded coordinates; missing slots break lines; positive and negative bars meet the zero baseline. | Fake-canvas primitive/coordinate checks plus visible English/Korean UI inspection at the supported sizes. The sketch is not a native PowerPoint renderer. |

These contracts reduce particular information-loss failures. They do not establish source truth, completeness of selected rows, appropriate units or denominators, semantic correctness of arbitrary prose, causal relationships, or final Office layout. The chosen-prose check compares magnitudes; signs, units, subjects and the date/quarter exclusions remain separate. Scientific-literal support in the checker does not ensure that every fact producer retains a whole scientific literal; incomplete metadata can conservatively reject it. Manual content selections still lack a complete schema-drift contract. These claims require their own evidence.

## Licensing record

- Original fixes and this documentation → existing [MIT License](../LICENSE) → consistent with the project's current permissive distribution; preserve its copyright and permission notice.
- Literature → citation and short original summaries only → no paper code, figures or full text is incorporated into distributed project material. A citation does not grant rights to reuse those materials; obtain the applicable permission before copying them.
- Existing bundled components → their existing [third-party notices](../THIRD-PARTY-NOTICES.md) → no new runtime dependency or model is introduced by these methods. This focused review does not certify comprehensive legal clearance.
