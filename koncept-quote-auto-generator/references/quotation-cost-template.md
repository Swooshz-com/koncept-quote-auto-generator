---
title: Quotation Cost Template
kind: pricing_source
currency: SGD
rows: 103
priced_rows: 103
sections: 11
continuation_rows_folded: 36
section_header_rows_omitted: 11
---

# Quotation Cost Template

Authoritative pricing source for `scripts/generate_quote.py`.

## Template Schema

| Field | Meaning |
| --- | --- |
| Row | Stable template row number used for auditing and match reports. |
| Default Qty | Default quantity when the template includes one. |
| Item | Priceable item text. Continuation lines are joined with `<br>`. |
| Estimate | Stored estimate value from the template. |
| Cost | Internal base cost. |
| GST Multiplier | GST multiplier when the row includes GST. Blank means no multiplier. |
| Mark-Up | Mark-up multiplier. |
| Remarks | Source remark, supplier note, keyword, finish note, retained extra value, or continuation note. |

## Parser Rules

- Section headings start pricing sections.
- A table row is priceable when `Item` is present, `Cost` is greater than zero, and `Mark-Up` is greater than zero.
- Continuation text from blank-cost rows is folded into the nearest priced row above it.
- Section header rows are represented by `###` headings, not duplicated inside item tables.

## Template Summary

| Metric | Count |
| --- | ---: |
| Sections | 11 |
| Price rows | 103 |
| Priceable rows | 103 |
| Continuation rows folded | 36 |
| Section header rows represented as headings | 11 |

## Pricing Template By Section

### 1. Floor Design

Anchor: `1-floor-design`

| Metric | Value |
| --- | ---: |
| Row range | 4-14 |
| Price rows | 6 |

| Row | Default Qty | Item | Estimate | Cost | GST Multiplier | Mark-Up | Remarks |
| ---: | ---: | --- | ---: | ---: | ---: | ---: | --- |
| 4 |  | m2 needle punch carpet in           colour | 0 | 7 | 1.09 | 1.5 | needle punch |
| 6 |  | m2 needle velour carpet in           colour | 0 | 19 | 1.09 | 1.5 | velour |
| 8 |  | m2 100mm raised platfrom with aluminum edging | 0 | 40 | 1.09 | 1.5 | Platform ONLY |
| 10 |  | m2 white laminated flooring on raised platform | 0 | 35 | 1.09 | 1.5 | white laminate |
| 12 |  | m2 wood grain laminated flooring on raised platform | 0 | 50 | 1.09 | 1.5 | wood grain laminate |
| 14 |  | m2 2mm PVC flooring on raised platform | 0 | 40 | 1.09 | 1.5 | PVC |

### 2. Booth Structure

Anchor: `2-booth-structure`

| Metric | Value |
| --- | ---: |
| Row range | 19-77 |
| Price rows | 20 |

| Row | Default Qty | Item | Estimate | Cost | GST Multiplier | Mark-Up | Remarks |
| ---: | ---: | --- | ---: | ---: | ---: | ---: | --- |
| 19 |  | m length single side partition wall at height 2.4m<br>wooden construct in painted finished as per design proposal | 0 | 180 |  | 1.5 | Backwall or any partition<br>PAINTED |
| 22 |  | m length single side partition wall at height 2.4m<br>wooden construct in laminated finished as per design proposal | 0 | 280 |  | 1.5 | Backwall or any partition<br>LAMINATED |
| 25 |  | m length double side partition wall at height 2.4m<br>wooden construct in  painted finished as per design proposal | 0 | 360 |  | 1.5 | Double sided partition<br>PAINTED |
| 28 |  | m length double side partition wall at height 2.4m<br>wooden construct in  laminated finished as per design proposal | 0 | 560 |  | 1.5 | Double sided partition<br>LAMINATE |
| 31 |  | m length single side partition wall at height 4m<br>wooden construct in painted finished as per design proposal | 0 | 360 |  | 1.5 | Backwall or any partition<br>PAINTED |
| 34 |  | m length single side partition wall at height 4m<br>wooden construct in  laminated finished as per design proposal | 0 | 480 |  | 1.5 | Backwall or any partition<br>LAMINATED |
| 37 |  | m length double side partition wall at height 4m<br>wooden construct in painted finished as per design proposal | 0 | 580 |  | 1.5 | Double sided partition<br>PAINTED |
| 40 |  | m length double side partition wall at height 4m<br>wooden construct in  laminated finished as per design proposal | 0 | 850 |  | 1.5 | Double sided partition<br>LAMINATE |
| 44 |  | m length top fascia structure at height 3.99m<br>wooden construct in painted finished as per design proposal | 0 | 250 |  | 1.5 | Top Booth Structure<br>PAINTED |
| 47 |  | m length top fascia structure at height 3.99m<br>wooden construct in laminated finished as per design proposal | 0 | 450 |  | 1.5 | Top Booth Structure<br>LAMINATED |
| 50 |  | m length double side partition wall at height 2.5m for meeting room<br>wooden construct in painted finished as per design proposal | 0 | 360 |  | 1.5 | Partition for meeting room<br>PAINTED |
| 53 |  | m length double side partition wall at height 2.5m for meeting room<br>wooden construct in laminated finished as per design proposal | 0 | 560 |  | 1.5 | Partition for meeting room<br>LAMINATE |
| 56 |  | nos. wooden construct low parition ready in painted finished | 0 | 180 |  | 1.5 | Low dividing Partition<br>PAINTED |
| 59 |  | nos. wooden construct low parition ready in laminated finished | 0 | 280 |  | 1.5 | Low dividing Partition<br>LAMINATED |
| 62 |  | m length 1mW x 2.5mH Sytem Profile partition with clear acrylic panel | 0 | 185 |  | 1.5 | Acrylic System Partition. Extra calculations: 150+15+20<br>Octanorm |
| 65 |  | nos. vertical support pillars in painted finished | 0 | 450 |  | 1.5 | Support pillar<br>PAINTED |
| 68 |  | nos. vertical support pillars in laminated finished | 0 | 850 |  | 1.5 | Support pillar<br>LAMINATED |
| 71 |  | nos. planter box in painted finished | 0 | 350 |  | 1.5 | PLANT BOX<br>PAINTED |
| 74 |  | nos. planter box in laminated finished | 0 | 550 |  | 1.5 | PLANT BOX<br>PAINTED |
| 77 |  | m length x 2.5m height glass partition | 0 | 820 |  | 1.5 | GLASS PARTITION |

### 3. COUNTERS AND CABINETS

Anchor: `3-counters-and-cabinets`

| Metric | Value |
| --- | ---: |
| Row range | 83-101 |
| Price rows | 7 |

| Row | Default Qty | Item | Estimate | Cost | GST Multiplier | Mark-Up | Remarks |
| ---: | ---: | --- | ---: | ---: | ---: | ---: | --- |
| 83 |  | m length x 1m height x 0.5m Width lockable information counter<br>wooden construct in painted finished and laminated top as per design proposal | 0 | 800 |  | 1.5 | INFORMATION COUNTER<br>PAINTED |
| 86 |  | m length x 1m height x 0.5m Width lockable information counter<br>wooden construct in laminated finished as per design proposal | 0 | 1200 |  | 1.5 | INFORMATION COUNTER<br>FULL LAMINATED |
| 89 |  | nos. of 1m length x 1m height x 0.5m Width lockable counter<br>wooden construct in painted finished and laminated top as per design proposal | 0 | 800 |  | 1.5 | COUNTER<br>PAINTED |
| 92 |  | nos. of 1m length x 1m height x 0.5m Width lockable counter<br>wooden construct in laminated finished as per design proposal | 0 | 1200 |  | 1.5 | COUNTER<br>FULL LAMINATED |
| 95 |  | nos. of 1m length x 1m height x 0.5m Width lockable counter with glass display top<br>wooden construct in painted finished and laminated top as per design proposal | 0 | 1200 |  | 1.5 | COUNTER WITH GLASS TOP<br>PAINTED |
| 98 |  | nos. of 1m length x 1m height x 0.5m Width lockable counter with glass display top<br>wooden construct in laminated finished as per design proposal | 0 | 1400 |  | 1.5 | COUNTER<br>FULL LAMINATED |
| 101 |  | Professional Engineer Endorsement for structure above 4m | 0 | 800 | 1.09 | 1.5 | PE |

### 4. Hanging Structure

Anchor: `4-hanging-structure`

| Metric | Value |
| --- | ---: |
| Row range | 106-125 |
| Price rows | 8 |

| Row | Default Qty | Item | Estimate | Cost | GST Multiplier | Mark-Up | Remarks |
| ---: | ---: | --- | ---: | ---: | ---: | ---: | --- |
| 106 |  | m rental of 300mm x 300mm Aluminium Box Truss | 0 | 90 | 1.09 | 1.5 | TRUSS |
| 108 |  | nos. rigging point for Overhead Structure or Aluminium Box Truss<br>• Prices are not inclusive of truss | 0 | 300 | 1.09 | 1.5 | RIGGING POINT |
| 111 |  | nos. of Manual Chain Hoist | 0 | 500 | 1.09 | 1.5 | MANUAL HOIST |
| 113 | 1 | Lot. rental of Boom Lift for Rigging (Mandatory charge per booth) | 817.5 | 500 | 1.09 | 1.5 | BOOM LIFT ONLY FOR EXPO |
| 115 |  | m run of hanging structure x 1m height<br>wooden construct in painted finished as per design proposal | 0 | 300 | 1.09 | 1.5 | Wooden hanging structure<br>PAINTED |
| 118 |  | m run of hanging structure x 1m height<br>wooden construct in laminated finished as per design proposal | 0 | 450 | 1.09 | 1.5 | Wooden hanging structure<br>LAMINATE |
| 121 |  | m run of hanging structure x 1m height<br>aluminum profile construct in fabric finished as per design proposal<br>Including fabric print | 0 | 100 | 1.09 | 1.5 | TUBE with PRINTED FABRIC |
| 125 |  | Professional Engineer Endorsement for hanging | 0 | 800 | 1.09 | 1.5 | PE. Extra values: 780; 68. Extra calculations: 90+690; 17*4<br>Extra values: 1000; 14.705882352941176. Extra calculations: 1000/68 |

### 5. Electrical Fittings ( Excluding connection fees by Organiser)

Anchor: `5-electrical-fittings-excluding-connection-fees-by-organiser`

| Metric | Value |
| --- | ---: |
| Row range | 130-150 |
| Price rows | 11 |

| Row | Default Qty | Item | Estimate | Cost | GST Multiplier | Mark-Up | Remarks |
| ---: | ---: | --- | ---: | ---: | ---: | ---: | --- |
| 130 |  | nos. 10W LED Spotlight | 0 | 30 | 1.09 | 1.5 | SPOTLIGHT |
| 132 |  | nos. 10W Arm LED Spotlight | 0 | 35 | 1.09 | 1.5 | ARM SPOTLIGHT |
| 134 |  | nos. 70W LED Arm Floodlight | 0 | 70 | 1.09 | 1.5 | ARM Flood light |
| 136 |  | nos. 150W HQI Floodlight | 0 | 70 | 1.09 | 1.5 | HQI Floodlight |
| 138 |  | nos. LED recess dowlight 3" | 0 | 30 | 1.09 | 1.5 | Downlight |
| 140 |  | nos. LED recess dowlight 6" | 0 | 35 | 1.09 | 1.5 | Downlight |
| 142 |  | nos.13Amp/230V SP 50Hz AC Socket (Max 800W) (Not for lighting use) | 0 | 95 | 1.09 | 1.5 | POWERPOINT |
| 144 |  | nos. 13Amp/230V SP 50Hz AC Socket with 24hrs Supply (Max 800W) (Not for lighting use) | 0 | 180 | 1.09 | 1.5 | POWERPOINT |
| 146 |  | m. run LED strip light for coves | 0 | 28 | 1.09 | 1.5 | LED STRIPS |
| 148 |  | m. run LED bar for backlit graphics | 0 | 40 | 1.09 | 1.5 | LED BAR |
| 150 |  | no. single line drawing for DB box | 0 | 400 | 1.09 | 1.5 | Single Line Drawing |

### 6. Furniture Rental

Anchor: `6-furniture-rental`

| Metric | Value |
| --- | ---: |
| Row range | 154-202 |
| Price rows | 25 |

| Row | Default Qty | Item | Estimate | Cost | GST Multiplier | Mark-Up | Remarks |
| ---: | ---: | --- | ---: | ---: | ---: | ---: | --- |
| 154 |  | nos. Bistro Chairs | 0 | 30 | 1.09 | 1.5 | Bistro Low Chair |
| 156 |  | nos. Eames Replica Chair (White) | 0 | 30 | 1.09 | 1.5 | Bistro Low Chair |
| 158 |  | nos. low Lidas Chair | 0 | 30 | 1.09 | 1.5 | Bistro Low Chair |
| 160 |  | nos. Chair With Cushion | 0 | 30 | 1.09 | 1.5 | Bistro Low Chair |
| 162 |  | nos. Eames-Style Replica Arm Chair (Black) | 0 | 50 | 1.09 | 1.5 | Bistro Low Chair |
| 164 |  | nos. seating chair (grey) | 0 | 25 | 1.09 | 1.5 | Seating chair |
| 166 |  | nos. Opel replica chairs (white) | 0 | 35 | 1.09 | 1.5 | Seating chair |
| 168 |  | nos. white folding chairs | 0 | 8 | 1.09 | 1.5 | Folding Chair |
| 170 |  | nos. Eames-Style Bar Stool (Replica) | 0 | 35 | 1.09 | 1.5 | Bar Stool |
| 172 |  | nos. Max White Bar Stool | 0 | 35 | 1.09 | 1.5 | Bar Stool |
| 174 |  | nos. Jay White Bar Stool with high back rest | 0 | 35 | 1.09 | 1.5 | Bar Stool |
| 176 |  | nos. Don White Bar Stool | 0 | 30 | 1.09 | 1.5 | Bar Stool |
| 178 |  | nos. Saffron Bar Stool (Black) | 0 | 35 | 1.09 | 1.5 | Bar Stool |
| 180 |  | nos. Roudy Bar Stool (Black) | 0 | 35 | 1.09 | 1.5 | Bar Stool |
| 182 |  | nos. Roudy Bar Stool (White) | 0 | 35 | 1.09 | 1.5 | Bar Stool |
| 184 |  | nos. Juliette Bar Stool | 0 | 35 | 1.09 | 1.5 | Bar Stool |
| 186 |  | nos. Conference Chair | 0 | 55 | 1.09 | 1.5 | Conference chairs |
| 188 |  | nos. Faux Leather Black Chair Twin Seating | 0 | 125 | 1.09 | 1.5 | SOFA chairs |
| 190 |  | nos. Modular Faux Leather White Sofa per unit | 0 | 75 | 1.09 | 1.5 | SOFA SEATS |
| 192 |  | nos. High Top Table White | 0 | 75 | 1.09 | 1.5 | BISTRO TABLE HIGH |
| 194 |  | nos. Aluminum Bistro Table (Round) | 0 | 75 | 1.09 | 1.5 | BISTRO TABLE HIGH |
| 196 |  | nos. Aluminum Bistro Table (Square) | 0 | 75 | 1.09 | 1.5 | BISTRO TABLE HIGH |
| 198 |  | nos. Aluminum Bistro Table (Square) | 0 | 75 | 1.09 | 1.5 | BISTRO TABLE HIGH |
| 200 |  | nos. White Tulip Table (60cm) | 0 | 45 | 1.09 | 1.5 | ROUNDTABLE LOW |
| 202 |  | nos. Round Glass Low Table (90cm) | 0 | 45 | 1.09 | 1.5 | GLASS ROUNDTABLE LOW |

### 7. Rental Items

Anchor: `7-rental-items`

| Metric | Value |
| --- | ---: |
| Row range | 208-220 |
| Price rows | 7 |

| Row | Default Qty | Item | Estimate | Cost | GST Multiplier | Mark-Up | Remarks |
| ---: | ---: | --- | ---: | ---: | ---: | ---: | --- |
| 208 |  | nos. Aluminum Foldable Brochure Stand | 0 | 55 | 1.09 | 1.5 | Brochure stand |
| 210 |  | nos. Aluminum Foldable Magazine Stand | 0 | 55 | 1.09 | 1.5 | Brochure stand |
| 212 |  | nos. Queue pole 1m (Silver) | 0 | 18 | 1.09 | 1.5 | Q POLES |
| 214 |  | nos. Queue pole 1m (Black) | 0 | 18 | 1.09 | 1.5 | Q POLES |
| 216 |  | nos. 420L Freezer (Mini Fridge) | 0 | 150 | 1.09 | 1.5 | MINI FRIDGE |
| 218 |  | nos. 550L Refrigerator | 0 | 150 | 1.09 | 1.5 | MINI FRIDGE |
| 220 |  | nos. Sink with Cabinet | 0 | 90 | 1.09 | 1.5 | SINK |

### 8. AV Equipment Rental Items

Anchor: `8-av-equipment-rental-items`

| Metric | Value |
| --- | ---: |
| Row range | 225-231 |
| Price rows | 4 |

| Row | Default Qty | Item | Estimate | Cost | GST Multiplier | Mark-Up | Remarks |
| ---: | ---: | --- | ---: | ---: | ---: | ---: | --- |
| 225 |  | nos. 42” LED TV Monitor (With Speaker – Full HD) | 0 | 300 | 1.09 | 1.5 | TV |
| 227 |  | nos. 55” LED TV Monitor (With Speaker – Full HD) | 0 | 450 | 1.09 | 1.5 | TV |
| 229 |  | nos. 65” LED TV Monitor (With Speaker – Full HD) | 0 | 650 | 1.09 | 1.5 | TV |
| 231 |  | nos. 85” LED TV Monitor (With Speaker – Full HD) | 0 | 1500 | 1.09 | 1.5 | TV |

### 9. Graphics

Anchor: `9-graphics`

| Metric | Value |
| --- | ---: |
| Row range | 238-263 |
| Price rows | 11 |

| Row | Default Qty | Item | Estimate | Cost | GST Multiplier | Mark-Up | Remarks |
| ---: | ---: | --- | ---: | ---: | ---: | ---: | --- |
| 238 |  | nos. 3D backlit logo | 0 | 350 | 1.09 | 1.5 | Backlit Logo |
| 240 |  | sets of 3D backlit lettering | 0 | 400 | 1.09 | 1.5 | Backlit Lettering<br>NOTE: Per set max 2m long |
| 243 |  | nos. 2D backlit logo | 0 | 200 | 1.09 | 1.5 | Backlit Lettering |
| 245 |  | nos. 3D vinyl logo on foam | 0 | 150 | 1.09 | 1.5 | 3D cut-out logo |
| 247 |  | sets of 3D vinyl lettering on foam | 0 | 160 | 1.09 | 1.5 | 3D cut-out lettering<br>NOTE: Per set max 2m long |
| 250 |  | nos. die-cut vinyl logo including lettering | 0 | 50 | 1.09 | 1.5 | 3D cut-out lettering<br>NOTE: Per set max 2m long |
| 253 |  | m2 of vinyl printed graphics | 0 | 40 | 1.09 | 1.5 | Printed Graphics on wall |
| 255 |  | m2 of vinyl printed graphics on flex foam | 0 | 48 | 1.09 | 1.5 | Printed Graphics on foam |
| 257 |  | nos. digital print graphic mounted directly onto system panels <br>(Size: 950mmL x 2340mmH) | 0 | 100 | 1.09 | 1.5 | For Shell Scheme |
| 260 |  | nos.seamless wall graphic on flex foam<br>(Size: 3000mmL x 2500mmH) | 0 | 360 | 1.09 | 1.5 | For Shell Scheme |
| 263 |  | nos. poster print on foam for exhibitor wall at approx. 1mL x 1.2mH | 0 | 60 | 1.09 | 1.5 | POSTER on foam |

### 10. Water Connection

Anchor: `10-water-connection`

| Metric | Value |
| --- | ---: |
| Row range | 268-270 |
| Price rows | 2 |

| Row | Default Qty | Item | Estimate | Cost | GST Multiplier | Mark-Up | Remarks |
| ---: | ---: | --- | ---: | ---: | ---: | ---: | --- |
| 268 |  | nos. water inlet and outlet | 0 | 720 | 1.09 | 1.2 | Water connection |
| 270 |  | nos. sink connection | 0 | 180 | 1.09 | 1.2 | Plumbling |

### 11. Coffee / Tea (Subject to approval by Venue owner and Organiser)

Anchor: `11-coffee-tea-subject-to-approval-by-venue-owner-and-organiser`

| Metric | Value |
| --- | ---: |
| Row range | 276-278 |
| Price rows | 2 |

| Row | Default Qty | Item | Estimate | Cost | GST Multiplier | Mark-Up | Remarks |
| ---: | ---: | --- | ---: | ---: | ---: | ---: | --- |
| 276 |  | Coffee/ Tea and supplies for 100 people per day | 0 | 150 | 1.09 | 1.5 | COFFEE PER DAY |
| 278 |  | Water cooler and heating equipment for 3 gallons | 0 | 250 | 1.09 | 1.5 | DRINKING WATER |
