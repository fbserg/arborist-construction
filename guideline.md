# Report Content Guidelines

This file defines the content rules, logic framework, and narrative templates for arborist reports. For technical workflow (reading, editing, packing `.docx` files), see `CLAUDE.md`.

---

### 1. Structural Layout & Formatting

Every assessment must follow this strict hierarchy:

*   **Markdown Data Table:** Every assessment begins with a table summarizing the tree data.
*   **Tree Condition & Context:** A narrative paragraph describing the tree's health and location.
*   **Impact Analysis:** A paragraph detailing the distance of work and TPZ encroachment.
*   **Construction Methodology:** Technical details of the excavation depth and method.
*   **Root Expectation & Mitigation:** What roots will be found and how to handle them.
*   **Prognosis:** Final sentence on survivability using species tolerance logic.

**The Data Table Schema**
The table must always include the following columns (omit Crown Dia if data is unavailable):
| TREE # | Species | DBH (cm) | Condition Rating | Comments | Ownership Category | Direction (Injury/Remove) | TPZ (m) | % Encroachment | Permit Requirement (Yes/No) |

### 2. Logic Framework: Impact Profiles

The LLM must categorize the user's proposed work into one of the following Standard Impact Profiles.

**Profile A: Surface Hardscaping (Driveways, Walkways, Patios, Sports Courts)**
*   **Excavation Depth:** Strict limit of 4 to 6 inches (10-15cm).
*   **Base Material:** 3-inch gravel/decomposed granite base + 2-3 inch paver/asphalt layer.
*   **Root Expectation:** Minor to moderate amount of small/fibrous roots (<2cm). Large roots (>5cm) unlikely.
*   **Impact Rating:** Minor to Moderate.
*   **Key Phrase:** "Excavation deeper than [4 or 6] inches is not allowed."

**Profile B: Isolated Vertical Structures (Fence Posts, Deck Posts, Pergolas)**
*   **Excavation Depth:** Up to 48 inches (1.2m) or 15 inches for minor decks.
*   **Methodology:** 4x4 posts in Sono-tubes (usually 12-inch diameter).
*   **Root Expectation:** Variable.
*   **Mitigation Strategy:** If a root >5cm is hit, move the post hole. Do not cut the root.
*   **Impact Rating:** Minor.

**Profile C: Foundations & Structures (Houses, Garages, Pools, Retaining Walls)**
*   **Excavation Depth:** Over 6 feet (Deep excavation).
*   **Mitigation (The "No Overdig" Rule):** Traditional construction requires a working space ("overdig") of 1 meter beyond the foundation wall.
*   **Required Recommendation:** "Vertical shoring (blind-side forming) is required along the limit of excavation to eliminate the need for a 1m overdig within the TPZ."
*   **Impact Rating:** Moderate to Severe.
*   **Root Expectation:** Total root loss within the excavation footprint.

**Profile D: Demolition (Shed, Existing Driveway/Patio Removal)**
*   **Methodology:** "Demolished by hand." No heavy machinery in TPZ.
*   **Concrete Removal:** Break concrete with hand tools inside TPZ; jackhammers allowed only outside TPZ.
*   **Impact Rating:** Minor (often beneficial due to increased permeability).

**Profile E: Trenching (Utilities, Water Service, Electrical)**
*   **Excavation Depth:** Up to 48 inches.
*   **Methodology:** Root Sensitive Excavation (RSE). Hand digging, AirSpade, or HydroVac.
*   **Mitigation:** Tunneling under roots >5cm.

### 3. Logic Framework: Species Tolerance Modifier

Before assigning an Impact Rating, the LLM must assess the species' tolerance to root disturbance.

*   **Tolerance Categories:**
    *   **Poor:** Oaks (*Quercus*), Beech (*Fagus*), Birch (*Betula*), Sugar Maple (*Acer saccharum*).
    *   **Moderate:** Norway Maple (*Acer platanoides*), Spruce (*Picea*), Pine (*Pinus*).
    *   **Good:** Willow (*Salix*), Poplar (*Populus*), Elm (*Ulmus*), Ash (*Fraxinus*).

*   **The Adjustment Rule:**
    *   If Species Tolerance is **Poor** AND Encroachment is **>15%**:
    *   **Action:** Upgrade Impact Rating to **Severe** and recommend heightened mitigation (e.g., AirSpade supervision required).

### 4. Narrative Generation Templates

**Template 1: Injury Assessment (Standard)**

> **[Para 1: Context]** Tree #[Number] is a [DBH]cm [Age Class: semi-mature/mature] [Species] in [Ownership] ownership. The subject tree is in [Condition] condition. [Insert specific flaws from comments]. These flaws are not of immediate concern.
>
> **[Para 2: The Intrusion]** The subject tree will require [Impact Rating: minor/moderate] TPZ encroachment and root injuries to allow for the construction of the [Project]. Footprints of the [Project] are located approximately [Distance]m away from the trunk at the closest point, necessitating [Encroachment Amount]m encroachment on its [TPZ Size]m TPZ.
>
> **[Para 3: Methodology]** The [Project] will be constructed as a [permeable/non-permeable] surface requiring [Depth: 4-6 inches] excavation to allow for a [Base Depth] inch base and [Surface Depth] inch [Material] layer. Excavation deeper than [Depth] inches is not allowed.
>
> **[Para 4: Roots & Mitigation]** Footprints of the proposal overlap the TPZ by approximately [Overlap %], and at this distance, it is likely that [Root Size: small/medium] roots will be discovered. Roots larger than 5cm must be retained. If uncovered, the envelope should be backfilled and the footprint reduced. Pruning minor amounts of small roots is expected to have no effect on structural integrity.
>
> **[Para 5: Conclusion]** Provided the mitigation measures are followed, the proposed injury is within the species' tolerance levels and the tree is expected to remain viable.

**Template 2: Removal Assessment (Direct Conflict)**

> **[Para 1: Context]** Tree #[Number] is a [DBH]cm [Species] in [Ownership] ownership. The subject tree is in [Condition] condition. [List flaws].
>
> **[Para 2: Conflict]** The subject tree will require removal to allow for the construction of [Project]. It is located directly within the footprints (or "directly adjacent, less than 1m away") of the proposed structure.
>
> **[Para 3: Justification]** Due to its location, it is obstructing the space needed for construction/excavation. Retention is not possible, and pre-emptive removal is recommended.

**Template 3: Removal Assessment (Condition Based)**

> **[Para 1: Context]** Tree #[Number] is a [DBH]cm [Species]. The subject tree is in Poor botanical and structural condition.
>
> **[Para 2: Justification]** Significant flaws include [Deadwood/Cavities/Lean]. The tree is in advanced health decline. Regardless of the proposed construction, this tree is not suitable for long-term retention and poses a potential hazard. Removal is recommended due to condition.

### 5. Critical Constraints & Tone Check

*   **Terminology:** Never use "Invade." Use **"Encroachment"**, **"Incursion"**, or **"Trespass"**.
*   **Referencing:** Always use **"The subject tree"** or **"Tree #[X]"**. Never use "The tree in question."
*   **Tone:** The output must be clinical and objective.
    *   *Bad:* "Unfortunately, the tree must be removed." / "Sadly, the roots will be cut."
    *   *Good:* "The tree requires removal." / "Root pruning is required."
*   **Forbidden Depths:** Never allow excavation deeper than 6 inches for driveways/walkways/patios within a TPZ.
*   **Machine Limits:** Never allow heavy machinery within the TPZ for demolition; specify "Hand tools" or "Non-vibrating tools."

### 6. Logic for Recommendations (New Projects)

2.  **Calculate Encroachment:** TPZ - Distance to work. Example: 4.0m - 2.0m = 2.0m encroachment.
3.  **Check Tolerance:** Oak = **Poor Tolerance**.
5.  **Select Profile:** Garage = Profile C (Deep excavation).
6.  **Apply Mitigation:** "Vertical shoring / blind-side forming required."

### 7. Post-Removal/Conclusion Notes

The LLM should append specific notes based on the Municipality or Ownership:
