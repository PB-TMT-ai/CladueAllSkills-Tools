# Workflow 04: Team Assignment

## Overview
Assign responsible teams to activities using flowchart analysis and inference.

**Critical Rule**: NEVER GUESS - Always ASK if ambiguous

---

## Teams Reference

```
1. Sales - Blue boxes in flowchart
   • Customer interaction, opportunity creation
   • Systems: Distributor Portal, SFDC

2. Planning - Green boxes
   • Inventory analysis, plant coordination
   • Systems: Excel, Planning Systems, ERP

3. Biz-ops - Pink boxes
   • Order execution, system operations
   • Systems: SFDC, ERP, Email, Zoho

4. JOTS - Yellow boxes
   • Transportation, vehicle planning
   • Systems: Zoho, Internal Systems

5. Plant Operations - Purple boxes
   • Dispatch, invoicing, GRN
   • Systems: ERP, Plant Portal, IRP
```

---

## Step 1: Flowchart Analysis (2-3 min)

### 1.1 Match Activities to Flowchart
```
FOR EACH activity:
  1. Search flowchart for matching text
  2. Identify box color
  3. Map color → Team
  4. Confidence level: High/Medium/Low

HIGH CONFIDENCE:
  → Exact text match + clear color
  
MEDIUM CONFIDENCE:
  → Partial match or faded color
  
LOW CONFIDENCE:
  → No match or ambiguous
```

### 1.2 Document Mappings
```
OUTPUT FORMAT:
| Activity | Flowchart Box | Color | Team | Confidence |
|----------|---------------|-------|------|------------|
| 1.a      | "Login Portal"| Blue  | Sales| High       |
| 2.b      | [Not found]   | N/A   | ?    | Low        |
```

---

## Step 2: Inference Logic (3-5 min)

### 2.1 Keyword Analysis
```
SEARCH FOR:
• "Sales creates" → Sales
• "Planning analyzes" → Planning
• "Biz-ops coordinates" → Biz-ops
• "JOTS receives" → JOTS
• "Plant generates" → Plant Operations

ACTION: Scan activity description for keywords
```

### 2.2 System-Based Inference
```
SYSTEM → TEAM MAPPING:
• Distributor Portal → Sales
• SFDC (opportunities) → Sales OR Biz-ops
• Excel/Planning Systems → Planning
• ERP (order operations) → Biz-ops OR Plant
• Zoho → Biz-ops OR JOTS
• Plant Portal → Plant Operations
• IRP → Plant Operations

RULE: Use system + activity context together
```

### 2.3 Process Position
```
WORKFLOW POSITION:
• Pre-order activities → Usually Sales
• Order processing → Usually Biz-ops
• Transportation → Usually JOTS
• Dispatch/Invoicing → Usually Plant

CAUTION: Not always reliable alone
```

---

## Step 3: Flagging Ambiguous Cases (1-2 min)

### 3.1 Identify Ambiguous Activities
```
FLAG IF:
• No flowchart match
• Multiple candidate teams
• Conflicting indicators
• User explicitly questions

CREATE CLARIFICATION LIST:
1. Activity [X]: [Candidates + Reasoning]
2. Activity [Y]: [Candidates + Reasoning]
```

### 3.2 Present to User
```
Template:
"🔍 Team Clarification Needed

Activity 1.f: Freight Approval Request
Context: Submit freight cost to distributor
Candidates:
  • Sales: Creates and submits opportunity
  • Planning: Calculates freight costs

Which team submits for approval?

Activity 2.c: Order Number Generation
Context: Generate system order number
Candidates:
  • Biz-ops: System operations
  • Plant Operations: Final order processing

Which team generates order number?

Please specify teams for each."
```

---

## Step 4: Apply Assignments (1 min)

### 4.1 Update SOP Table
```
FOR EACH activity:
  1. Fill Team column with assigned team
  2. Verify no "TBD" or "Unknown" values
  3. Check consistency with related activities
  4. Ensure same activity doesn't switch teams
```

### 4.2 Final Verification
```
CHECKS:
[ ] All activities have team assigned
[ ] Teams match approved list
[ ] No typos in team names
[ ] Consistent across document
[ ] User confirmed all flagged items
```

---

## Decision Tree

```
START: Activity needs team assignment
  ↓
Flowchart has matching box?
  YES → Color clear? → YES → ASSIGN team → DONE
         ↓ NO
         Add to clarification list → ASK USER
  ↓ NO
Activity has team keyword?
  YES → Single candidate? → YES → ASSIGN team → DONE
         ↓ NO
         Add to clarification list → ASK USER
  ↓ NO
System name implies team?
  YES → Single candidate? → YES → ASSIGN team → DONE
         ↓ NO
         Add to clarification list → ASK USER
  ↓ NO
Process position suggests team?
  YES → High confidence? → YES → ASSIGN team → DONE
         ↓ NO
         Add to clarification list → ASK USER
  ↓ NO
NO INDICATORS → ASK USER
```

---

## Common Patterns

### Pattern 1: Sequential Handoff
```
EXAMPLE:
• 1.e: Sales submits opportunity → Sales
• 1.f: Freight approval → Sales (still owns process)
• 2.a: Planning receives approved order → Planning

RULE: Team owns until explicit handoff
```

### Pattern 2: System-Specific Operations
```
EXAMPLE:
• SFDC opportunity creation → Sales
• SFDC order operations → Biz-ops
• ERP inventory check → Planning
• ERP order confirmation → Biz-ops
• ERP invoice generation → Plant Operations

RULE: Same system, different team based on activity type
```

### Pattern 3: Approval Chains
```
EXAMPLE:
• Sales submits for approval → Sales
• RSM reviews and approves → Sales (RSM is Sales team)
• Order proceeds to Biz-ops → Biz-ops

RULE: Approval stays with submitting team
```

---

## Edge Cases

### EC1: Multi-Team Activity
```
IF activity involves 2+ teams:
  → Assign to PRIMARY responsible team
  → Note other teams in Steps column

EXAMPLE:
| 4.d | Coordinate with JOTS | Planning | Email/JOTS |
(Planning leads, JOTS participates)
```

### EC2: Customer-Facing Activities
```
IF customer interaction:
  → Usually Sales
  → Except post-dispatch customer support (Plant)

VERIFY with user if unclear
```

### EC3: System Administration
```
IF activity is system admin task:
  → Check who has system access
  → Often Biz-ops for operational systems
  → Plant Operations for plant-specific systems
```

---

## Success Criteria

Team assignment complete when:
1. ✅ All activities assigned
2. ✅ No "TBD" or "Unknown"
3. ✅ User confirmed ambiguous cases
4. ✅ Teams match approved list
5. ✅ Consistency verified
