---
title: Quotation Cost Template
kind: rag_pricing_source
currency: SGD
---

# Quotation Cost Template — Clean RAG Pricing Source

This is the retrieval-friendly quotation pricing source. It keeps only quote-relevant pricing knowledge and removes spreadsheet audit noise.

## Pricing calculation rules

- Currency: SGD.
- Internal cost is the base pre-GST cost per item or unit.
- Mark-up multiplier is applied to the internal cost for customer-facing quotation line amounts.
- Customer-facing quotation line amount = quantity × internal cost × mark-up multiplier.
- GST is shown separately in the quotation totals block, not added into each customer-facing line amount.
- GST multiplier is retained from the source template for internal/default quote amount references.
- Source-template default quote amount = default quantity × internal cost × GST multiplier × mark-up multiplier when a GST multiplier is shown.
- When no default quantity is shown, use the quantity from the quotation request.
- Default quote amount is shown only when the source template provides a default quantity and computed amount.

## Pricing items

### Floor Design

- **Item:** m2 needle punch carpet in colour.
  - **Internal cost:** SGD 7.
  - **GST multiplier:** 1.09.
  - **Mark-up multiplier:** 1.5.
  - **Remarks / search terms:** needle punch.

- **Item:** m2 needle velour carpet in colour.
  - **Internal cost:** SGD 19.
  - **GST multiplier:** 1.09.
  - **Mark-up multiplier:** 1.5.
  - **Remarks / search terms:** velour.

- **Item:** m2 100mm raised platfrom with aluminum edging.
  - **Internal cost:** SGD 40.
  - **GST multiplier:** 1.09.
  - **Mark-up multiplier:** 1.5.
  - **Remarks / search terms:** Platform ONLY.

- **Item:** m2 white laminated flooring on raised platform.
  - **Internal cost:** SGD 35.
  - **GST multiplier:** 1.09.
  - **Mark-up multiplier:** 1.5.
  - **Remarks / search terms:** white laminate.

- **Item:** m2 wood grain laminated flooring on raised platform.
  - **Internal cost:** SGD 50.
  - **GST multiplier:** 1.09.
  - **Mark-up multiplier:** 1.5.
  - **Remarks / search terms:** wood grain laminate.

- **Item:** m2 2mm PVC flooring on raised platform.
  - **Internal cost:** SGD 40.
  - **GST multiplier:** 1.09.
  - **Mark-up multiplier:** 1.5.
  - **Remarks / search terms:** PVC.

### Booth Structure

- **Item:** m length single side partition wall at height 2.4m; wooden construct in painted finished as per design proposal.
  - **Internal cost:** SGD 180.
  - **GST multiplier:** Not applied in source template.
  - **Mark-up multiplier:** 1.5.
  - **Remarks / search terms:** Backwall or any partition; PAINTED.

- **Item:** m length single side partition wall at height 2.4m; wooden construct in laminated finished as per design proposal.
  - **Internal cost:** SGD 280.
  - **GST multiplier:** Not applied in source template.
  - **Mark-up multiplier:** 1.5.
  - **Remarks / search terms:** Backwall or any partition; LAMINATED.

- **Item:** m length double side partition wall at height 2.4m; wooden construct in painted finished as per design proposal.
  - **Internal cost:** SGD 360.
  - **GST multiplier:** Not applied in source template.
  - **Mark-up multiplier:** 1.5.
  - **Remarks / search terms:** Double sided partition; PAINTED.

- **Item:** m length double side partition wall at height 2.4m; wooden construct in laminated finished as per design proposal.
  - **Internal cost:** SGD 560.
  - **GST multiplier:** Not applied in source template.
  - **Mark-up multiplier:** 1.5.
  - **Remarks / search terms:** Double sided partition; LAMINATE.

- **Item:** m length single side partition wall at height 4m; wooden construct in painted finished as per design proposal.
  - **Internal cost:** SGD 360.
  - **GST multiplier:** Not applied in source template.
  - **Mark-up multiplier:** 1.5.
  - **Remarks / search terms:** Backwall or any partition; PAINTED.

- **Item:** m length single side partition wall at height 4m; wooden construct in laminated finished as per design proposal.
  - **Internal cost:** SGD 480.
  - **GST multiplier:** Not applied in source template.
  - **Mark-up multiplier:** 1.5.
  - **Remarks / search terms:** Backwall or any partition; LAMINATED.

- **Item:** m length double side partition wall at height 4m; wooden construct in painted finished as per design proposal.
  - **Internal cost:** SGD 580.
  - **GST multiplier:** Not applied in source template.
  - **Mark-up multiplier:** 1.5.
  - **Remarks / search terms:** Double sided partition; PAINTED.

- **Item:** m length double side partition wall at height 4m; wooden construct in laminated finished as per design proposal.
  - **Internal cost:** SGD 850.
  - **GST multiplier:** Not applied in source template.
  - **Mark-up multiplier:** 1.5.
  - **Remarks / search terms:** Double sided partition; LAMINATE.

- **Item:** m length top fascia structure at height 3.99m; wooden construct in painted finished as per design proposal.
  - **Internal cost:** SGD 250.
  - **GST multiplier:** Not applied in source template.
  - **Mark-up multiplier:** 1.5.
  - **Remarks / search terms:** Top Booth Structure; PAINTED.

- **Item:** m length top fascia structure at height 3.99m; wooden construct in laminated finished as per design proposal.
  - **Internal cost:** SGD 450.
  - **GST multiplier:** Not applied in source template.
  - **Mark-up multiplier:** 1.5.
  - **Remarks / search terms:** Top Booth Structure; LAMINATED.

- **Item:** m length double side partition wall at height 2.5m for meeting room; wooden construct in painted finished as per design proposal.
  - **Internal cost:** SGD 360.
  - **GST multiplier:** Not applied in source template.
  - **Mark-up multiplier:** 1.5.
  - **Remarks / search terms:** Partition for meeting room; PAINTED.

- **Item:** m length double side partition wall at height 2.5m for meeting room; wooden construct in laminated finished as per design proposal.
  - **Internal cost:** SGD 560.
  - **GST multiplier:** Not applied in source template.
  - **Mark-up multiplier:** 1.5.
  - **Remarks / search terms:** Partition for meeting room; LAMINATE.

- **Item:** nos. wooden construct low parition ready in painted finished.
  - **Internal cost:** SGD 180.
  - **GST multiplier:** Not applied in source template.
  - **Mark-up multiplier:** 1.5.
  - **Remarks / search terms:** Low dividing Partition; PAINTED.

- **Item:** nos. wooden construct low parition ready in laminated finished.
  - **Internal cost:** SGD 280.
  - **GST multiplier:** Not applied in source template.
  - **Mark-up multiplier:** 1.5.
  - **Remarks / search terms:** Low dividing Partition; LAMINATED.

- **Item:** m length 1mW x 2.5mH Sytem Profile partition with clear acrylic panel.
  - **Internal cost:** SGD 185.
  - **GST multiplier:** Not applied in source template.
  - **Mark-up multiplier:** 1.5.
  - **Remarks / search terms:** Acrylic System Partition. Extra calculations: 150+15+20; Octanorm.

- **Item:** nos. vertical support pillars in painted finished.
  - **Internal cost:** SGD 450.
  - **GST multiplier:** Not applied in source template.
  - **Mark-up multiplier:** 1.5.
  - **Remarks / search terms:** Support pillar; PAINTED.

- **Item:** nos. vertical support pillars in laminated finished.
  - **Internal cost:** SGD 850.
  - **GST multiplier:** Not applied in source template.
  - **Mark-up multiplier:** 1.5.
  - **Remarks / search terms:** Support pillar; LAMINATED.

- **Item:** nos. planter box in painted finished.
  - **Internal cost:** SGD 350.
  - **GST multiplier:** Not applied in source template.
  - **Mark-up multiplier:** 1.5.
  - **Remarks / search terms:** PLANT BOX; PAINTED.

- **Item:** nos. planter box in laminated finished.
  - **Internal cost:** SGD 550.
  - **GST multiplier:** Not applied in source template.
  - **Mark-up multiplier:** 1.5.
  - **Remarks / search terms:** PLANT BOX; PAINTED.

- **Item:** m length x 2.5m height glass partition.
  - **Internal cost:** SGD 820.
  - **GST multiplier:** Not applied in source template.
  - **Mark-up multiplier:** 1.5.
  - **Remarks / search terms:** GLASS PARTITION.

### COUNTERS AND CABINETS

- **Item:** m length x 1m height x 0.5m Width lockable information counter; wooden construct in painted finished and laminated top as per design proposal.
  - **Internal cost:** SGD 800.
  - **GST multiplier:** Not applied in source template.
  - **Mark-up multiplier:** 1.5.
  - **Remarks / search terms:** INFORMATION COUNTER; PAINTED.

- **Item:** m length x 1m height x 0.5m Width lockable information counter; wooden construct in laminated finished as per design proposal.
  - **Internal cost:** SGD 1200.
  - **GST multiplier:** Not applied in source template.
  - **Mark-up multiplier:** 1.5.
  - **Remarks / search terms:** INFORMATION COUNTER; FULL LAMINATED.

- **Item:** nos. of 1m length x 1m height x 0.5m Width lockable counter; wooden construct in painted finished and laminated top as per design proposal.
  - **Internal cost:** SGD 800.
  - **GST multiplier:** Not applied in source template.
  - **Mark-up multiplier:** 1.5.
  - **Remarks / search terms:** COUNTER; PAINTED.

- **Item:** nos. of 1m length x 1m height x 0.5m Width lockable counter; wooden construct in laminated finished as per design proposal.
  - **Internal cost:** SGD 1200.
  - **GST multiplier:** Not applied in source template.
  - **Mark-up multiplier:** 1.5.
  - **Remarks / search terms:** COUNTER; FULL LAMINATED.

- **Item:** nos. of 1m length x 1m height x 0.5m Width lockable counter with glass display top; wooden construct in painted finished and laminated top as per design proposal.
  - **Internal cost:** SGD 1200.
  - **GST multiplier:** Not applied in source template.
  - **Mark-up multiplier:** 1.5.
  - **Remarks / search terms:** COUNTER WITH GLASS TOP; PAINTED.

- **Item:** nos. of 1m length x 1m height x 0.5m Width lockable counter with glass display top; wooden construct in laminated finished as per design proposal.
  - **Internal cost:** SGD 1400.
  - **GST multiplier:** Not applied in source template.
  - **Mark-up multiplier:** 1.5.
  - **Remarks / search terms:** COUNTER; FULL LAMINATED.

- **Item:** Professional Engineer Endorsement for structure above 4m.
  - **Internal cost:** SGD 800.
  - **GST multiplier:** 1.09.
  - **Mark-up multiplier:** 1.5.
  - **Remarks / search terms:** PE.

### Hanging Structure

- **Item:** m rental of 300mm x 300mm Aluminium Box Truss.
  - **Internal cost:** SGD 90.
  - **GST multiplier:** 1.09.
  - **Mark-up multiplier:** 1.5.
  - **Remarks / search terms:** TRUSS.

- **Item:** nos. rigging point for Overhead Structure or Aluminium Box Truss; - Prices are not inclusive of truss.
  - **Internal cost:** SGD 300.
  - **GST multiplier:** 1.09.
  - **Mark-up multiplier:** 1.5.
  - **Remarks / search terms:** RIGGING POINT.

- **Item:** nos. of Manual Chain Hoist.
  - **Internal cost:** SGD 500.
  - **GST multiplier:** 1.09.
  - **Mark-up multiplier:** 1.5.
  - **Remarks / search terms:** MANUAL HOIST.

- **Item:** Lot. rental of Boom Lift for Rigging (Mandatory charge per booth).
  - **Default quantity:** 1.
  - **Default quote amount:** SGD 817.5.
  - **Internal cost:** SGD 500.
  - **GST multiplier:** 1.09.
  - **Mark-up multiplier:** 1.5.
  - **Remarks / search terms:** BOOM LIFT ONLY FOR EXPO.

- **Item:** m run of hanging structure x 1m height; wooden construct in painted finished as per design proposal.
  - **Internal cost:** SGD 300.
  - **GST multiplier:** 1.09.
  - **Mark-up multiplier:** 1.5.
  - **Remarks / search terms:** Wooden hanging structure; PAINTED.

- **Item:** m run of hanging structure x 1m height; wooden construct in laminated finished as per design proposal.
  - **Internal cost:** SGD 450.
  - **GST multiplier:** 1.09.
  - **Mark-up multiplier:** 1.5.
  - **Remarks / search terms:** Wooden hanging structure; LAMINATE.

- **Item:** m run of hanging structure x 1m height; aluminum profile construct in fabric finished as per design proposal; Including fabric print.
  - **Internal cost:** SGD 100.
  - **GST multiplier:** 1.09.
  - **Mark-up multiplier:** 1.5.
  - **Remarks / search terms:** TUBE with PRINTED FABRIC.

- **Item:** Professional Engineer Endorsement for hanging.
  - **Internal cost:** SGD 800.
  - **GST multiplier:** 1.09.
  - **Mark-up multiplier:** 1.5.
  - **Remarks / search terms:** PE. Extra values: 780; 68. Extra calculations: 90+690; 17*4; Extra values: 1000; 14.705882352941176. Extra calculations: 1000/68.

### Electrical Fittings ( Excluding connection fees by Organiser)

- **Item:** nos. 10W LED Spotlight.
  - **Internal cost:** SGD 30.
  - **GST multiplier:** 1.09.
  - **Mark-up multiplier:** 1.5.
  - **Remarks / search terms:** SPOTLIGHT.

- **Item:** nos. 10W Arm LED Spotlight.
  - **Internal cost:** SGD 35.
  - **GST multiplier:** 1.09.
  - **Mark-up multiplier:** 1.5.
  - **Remarks / search terms:** ARM SPOTLIGHT.

- **Item:** nos. 70W LED Arm Floodlight.
  - **Internal cost:** SGD 70.
  - **GST multiplier:** 1.09.
  - **Mark-up multiplier:** 1.5.
  - **Remarks / search terms:** ARM Flood light.

- **Item:** nos. 150W HQI Floodlight.
  - **Internal cost:** SGD 70.
  - **GST multiplier:** 1.09.
  - **Mark-up multiplier:** 1.5.
  - **Remarks / search terms:** HQI Floodlight.

- **Item:** nos. LED recess dowlight 3".
  - **Internal cost:** SGD 30.
  - **GST multiplier:** 1.09.
  - **Mark-up multiplier:** 1.5.
  - **Remarks / search terms:** Downlight.

- **Item:** nos. LED recess dowlight 6".
  - **Internal cost:** SGD 35.
  - **GST multiplier:** 1.09.
  - **Mark-up multiplier:** 1.5.
  - **Remarks / search terms:** Downlight.

- **Item:** nos.13Amp/230V SP 50Hz AC Socket (Max 800W) (Not for lighting use).
  - **Internal cost:** SGD 95.
  - **GST multiplier:** 1.09.
  - **Mark-up multiplier:** 1.5.
  - **Remarks / search terms:** POWERPOINT.

- **Item:** nos. 13Amp/230V SP 50Hz AC Socket with 24hrs Supply (Max 800W) (Not for lighting use).
  - **Internal cost:** SGD 180.
  - **GST multiplier:** 1.09.
  - **Mark-up multiplier:** 1.5.
  - **Remarks / search terms:** POWERPOINT.

- **Item:** m. run LED strip light for coves.
  - **Internal cost:** SGD 28.
  - **GST multiplier:** 1.09.
  - **Mark-up multiplier:** 1.5.
  - **Remarks / search terms:** LED STRIPS.

- **Item:** m. run LED bar for backlit graphics.
  - **Internal cost:** SGD 40.
  - **GST multiplier:** 1.09.
  - **Mark-up multiplier:** 1.5.
  - **Remarks / search terms:** LED BAR.

- **Item:** no. single line drawing for DB box.
  - **Internal cost:** SGD 400.
  - **GST multiplier:** 1.09.
  - **Mark-up multiplier:** 1.5.
  - **Remarks / search terms:** Single Line Drawing.

### Furniture Rental

- **Item:** nos. Bistro Chairs.
  - **Internal cost:** SGD 30.
  - **GST multiplier:** 1.09.
  - **Mark-up multiplier:** 1.5.
  - **Remarks / search terms:** Bistro Low Chair.

- **Item:** nos. Eames Replica Chair (White).
  - **Internal cost:** SGD 30.
  - **GST multiplier:** 1.09.
  - **Mark-up multiplier:** 1.5.
  - **Remarks / search terms:** Bistro Low Chair.

- **Item:** nos. low Lidas Chair.
  - **Internal cost:** SGD 30.
  - **GST multiplier:** 1.09.
  - **Mark-up multiplier:** 1.5.
  - **Remarks / search terms:** Bistro Low Chair.

- **Item:** nos. Chair With Cushion.
  - **Internal cost:** SGD 30.
  - **GST multiplier:** 1.09.
  - **Mark-up multiplier:** 1.5.
  - **Remarks / search terms:** Bistro Low Chair.

- **Item:** nos. Eames-Style Replica Arm Chair (Black).
  - **Internal cost:** SGD 50.
  - **GST multiplier:** 1.09.
  - **Mark-up multiplier:** 1.5.
  - **Remarks / search terms:** Bistro Low Chair.

- **Item:** nos. seating chair (grey).
  - **Internal cost:** SGD 25.
  - **GST multiplier:** 1.09.
  - **Mark-up multiplier:** 1.5.
  - **Remarks / search terms:** Seating chair.

- **Item:** nos. Opel replica chairs (white).
  - **Internal cost:** SGD 35.
  - **GST multiplier:** 1.09.
  - **Mark-up multiplier:** 1.5.
  - **Remarks / search terms:** Seating chair.

- **Item:** nos. white folding chairs.
  - **Internal cost:** SGD 8.
  - **GST multiplier:** 1.09.
  - **Mark-up multiplier:** 1.5.
  - **Remarks / search terms:** Folding Chair.

- **Item:** nos. Eames-Style Bar Stool (Replica).
  - **Internal cost:** SGD 35.
  - **GST multiplier:** 1.09.
  - **Mark-up multiplier:** 1.5.
  - **Remarks / search terms:** Bar Stool.

- **Item:** nos. Max White Bar Stool.
  - **Internal cost:** SGD 35.
  - **GST multiplier:** 1.09.
  - **Mark-up multiplier:** 1.5.
  - **Remarks / search terms:** Bar Stool.

- **Item:** nos. Jay White Bar Stool with high back rest.
  - **Internal cost:** SGD 35.
  - **GST multiplier:** 1.09.
  - **Mark-up multiplier:** 1.5.
  - **Remarks / search terms:** Bar Stool.

- **Item:** nos. Don White Bar Stool.
  - **Internal cost:** SGD 30.
  - **GST multiplier:** 1.09.
  - **Mark-up multiplier:** 1.5.
  - **Remarks / search terms:** Bar Stool.

- **Item:** nos. Saffron Bar Stool (Black).
  - **Internal cost:** SGD 35.
  - **GST multiplier:** 1.09.
  - **Mark-up multiplier:** 1.5.
  - **Remarks / search terms:** Bar Stool.

- **Item:** nos. Roudy Bar Stool (Black).
  - **Internal cost:** SGD 35.
  - **GST multiplier:** 1.09.
  - **Mark-up multiplier:** 1.5.
  - **Remarks / search terms:** Bar Stool.

- **Item:** nos. Roudy Bar Stool (White).
  - **Internal cost:** SGD 35.
  - **GST multiplier:** 1.09.
  - **Mark-up multiplier:** 1.5.
  - **Remarks / search terms:** Bar Stool.

- **Item:** nos. Juliette Bar Stool.
  - **Internal cost:** SGD 35.
  - **GST multiplier:** 1.09.
  - **Mark-up multiplier:** 1.5.
  - **Remarks / search terms:** Bar Stool.

- **Item:** nos. Conference Chair.
  - **Internal cost:** SGD 55.
  - **GST multiplier:** 1.09.
  - **Mark-up multiplier:** 1.5.
  - **Remarks / search terms:** Conference chairs.

- **Item:** nos. Faux Leather Black Chair Twin Seating.
  - **Internal cost:** SGD 125.
  - **GST multiplier:** 1.09.
  - **Mark-up multiplier:** 1.5.
  - **Remarks / search terms:** SOFA chairs.

- **Item:** nos. Modular Faux Leather White Sofa per unit.
  - **Internal cost:** SGD 75.
  - **GST multiplier:** 1.09.
  - **Mark-up multiplier:** 1.5.
  - **Remarks / search terms:** SOFA SEATS.

- **Item:** nos. High Top Table White.
  - **Internal cost:** SGD 75.
  - **GST multiplier:** 1.09.
  - **Mark-up multiplier:** 1.5.
  - **Remarks / search terms:** BISTRO TABLE HIGH.

- **Item:** nos. Aluminum Bistro Table (Round).
  - **Internal cost:** SGD 75.
  - **GST multiplier:** 1.09.
  - **Mark-up multiplier:** 1.5.
  - **Remarks / search terms:** BISTRO TABLE HIGH.

- **Item:** nos. Aluminum Bistro Table (Square).
  - **Internal cost:** SGD 75.
  - **GST multiplier:** 1.09.
  - **Mark-up multiplier:** 1.5.
  - **Remarks / search terms:** BISTRO TABLE HIGH.

- **Item:** nos. Aluminum Bistro Table (Square).
  - **Internal cost:** SGD 75.
  - **GST multiplier:** 1.09.
  - **Mark-up multiplier:** 1.5.
  - **Remarks / search terms:** BISTRO TABLE HIGH.

- **Item:** nos. White Tulip Table (60cm).
  - **Internal cost:** SGD 45.
  - **GST multiplier:** 1.09.
  - **Mark-up multiplier:** 1.5.
  - **Remarks / search terms:** ROUNDTABLE LOW.

- **Item:** nos. Round Glass Low Table (90cm).
  - **Internal cost:** SGD 45.
  - **GST multiplier:** 1.09.
  - **Mark-up multiplier:** 1.5.
  - **Remarks / search terms:** GLASS ROUNDTABLE LOW.

### Rental Items

- **Item:** nos. Aluminum Foldable Brochure Stand.
  - **Internal cost:** SGD 55.
  - **GST multiplier:** 1.09.
  - **Mark-up multiplier:** 1.5.
  - **Remarks / search terms:** Brochure stand.

- **Item:** nos. Aluminum Foldable Magazine Stand.
  - **Internal cost:** SGD 55.
  - **GST multiplier:** 1.09.
  - **Mark-up multiplier:** 1.5.
  - **Remarks / search terms:** Brochure stand.

- **Item:** nos. Queue pole 1m (Silver).
  - **Internal cost:** SGD 18.
  - **GST multiplier:** 1.09.
  - **Mark-up multiplier:** 1.5.
  - **Remarks / search terms:** Q POLES.

- **Item:** nos. Queue pole 1m (Black).
  - **Internal cost:** SGD 18.
  - **GST multiplier:** 1.09.
  - **Mark-up multiplier:** 1.5.
  - **Remarks / search terms:** Q POLES.

- **Item:** nos. 420L Freezer (Mini Fridge).
  - **Internal cost:** SGD 150.
  - **GST multiplier:** 1.09.
  - **Mark-up multiplier:** 1.5.
  - **Remarks / search terms:** MINI FRIDGE.

- **Item:** nos. 550L Refrigerator.
  - **Internal cost:** SGD 150.
  - **GST multiplier:** 1.09.
  - **Mark-up multiplier:** 1.5.
  - **Remarks / search terms:** MINI FRIDGE.

- **Item:** nos. Sink with Cabinet.
  - **Internal cost:** SGD 90.
  - **GST multiplier:** 1.09.
  - **Mark-up multiplier:** 1.5.
  - **Remarks / search terms:** SINK.

### AV Equipment Rental Items

- **Item:** nos. 42" LED TV Monitor (With Speaker – Full HD).
  - **Internal cost:** SGD 300.
  - **GST multiplier:** 1.09.
  - **Mark-up multiplier:** 1.5.
  - **Remarks / search terms:** TV.

- **Item:** nos. 55" LED TV Monitor (With Speaker – Full HD).
  - **Internal cost:** SGD 450.
  - **GST multiplier:** 1.09.
  - **Mark-up multiplier:** 1.5.
  - **Remarks / search terms:** TV.

- **Item:** nos. 65" LED TV Monitor (With Speaker – Full HD).
  - **Internal cost:** SGD 650.
  - **GST multiplier:** 1.09.
  - **Mark-up multiplier:** 1.5.
  - **Remarks / search terms:** TV.

- **Item:** nos. 85" LED TV Monitor (With Speaker – Full HD).
  - **Internal cost:** SGD 1500.
  - **GST multiplier:** 1.09.
  - **Mark-up multiplier:** 1.5.
  - **Remarks / search terms:** TV.

### Graphics

- **Item:** nos. 3D backlit logo.
  - **Internal cost:** SGD 350.
  - **GST multiplier:** 1.09.
  - **Mark-up multiplier:** 1.5.
  - **Remarks / search terms:** Backlit Logo.

- **Item:** sets of 3D backlit lettering.
  - **Internal cost:** SGD 400.
  - **GST multiplier:** 1.09.
  - **Mark-up multiplier:** 1.5.
  - **Remarks / search terms:** Backlit Lettering; NOTE: Per set max 2m long.

- **Item:** nos. 2D backlit logo.
  - **Internal cost:** SGD 200.
  - **GST multiplier:** 1.09.
  - **Mark-up multiplier:** 1.5.
  - **Remarks / search terms:** Backlit Lettering.

- **Item:** nos. 3D vinyl logo on foam.
  - **Internal cost:** SGD 150.
  - **GST multiplier:** 1.09.
  - **Mark-up multiplier:** 1.5.
  - **Remarks / search terms:** 3D cut-out logo.

- **Item:** sets of 3D vinyl lettering on foam.
  - **Internal cost:** SGD 160.
  - **GST multiplier:** 1.09.
  - **Mark-up multiplier:** 1.5.
  - **Remarks / search terms:** 3D cut-out lettering; NOTE: Per set max 2m long.

- **Item:** nos. die-cut vinyl logo including lettering.
  - **Internal cost:** SGD 50.
  - **GST multiplier:** 1.09.
  - **Mark-up multiplier:** 1.5.
  - **Remarks / search terms:** 3D cut-out lettering; NOTE: Per set max 2m long.

- **Item:** m2 of vinyl printed graphics.
  - **Internal cost:** SGD 40.
  - **GST multiplier:** 1.09.
  - **Mark-up multiplier:** 1.5.
  - **Remarks / search terms:** Printed Graphics on wall.

- **Item:** m2 of vinyl printed graphics on flex foam.
  - **Internal cost:** SGD 48.
  - **GST multiplier:** 1.09.
  - **Mark-up multiplier:** 1.5.
  - **Remarks / search terms:** Printed Graphics on foam.

- **Item:** nos. digital print graphic mounted directly onto system panels; (Size: 950mmL x 2340mmH).
  - **Internal cost:** SGD 100.
  - **GST multiplier:** 1.09.
  - **Mark-up multiplier:** 1.5.
  - **Remarks / search terms:** For Shell Scheme.

- **Item:** nos.seamless wall graphic on flex foam; (Size: 3000mmL x 2500mmH).
  - **Internal cost:** SGD 360.
  - **GST multiplier:** 1.09.
  - **Mark-up multiplier:** 1.5.
  - **Remarks / search terms:** For Shell Scheme.

- **Item:** nos. poster print on foam for exhibitor wall at approx. 1mL x 1.2mH.
  - **Internal cost:** SGD 60.
  - **GST multiplier:** 1.09.
  - **Mark-up multiplier:** 1.5.
  - **Remarks / search terms:** POSTER on foam.

### Water Connection

- **Item:** nos. water inlet and outlet.
  - **Internal cost:** SGD 720.
  - **GST multiplier:** 1.09.
  - **Mark-up multiplier:** 1.2.
  - **Remarks / search terms:** Water connection.

- **Item:** nos. sink connection.
  - **Internal cost:** SGD 180.
  - **GST multiplier:** 1.09.
  - **Mark-up multiplier:** 1.2.
  - **Remarks / search terms:** Plumbling.

### Coffee / Tea (Subject to approval by Venue owner and Organiser)

- **Item:** Coffee/ Tea and supplies for 100 people per day.
  - **Internal cost:** SGD 150.
  - **GST multiplier:** 1.09.
  - **Mark-up multiplier:** 1.5.
  - **Remarks / search terms:** COFFEE PER DAY.

- **Item:** Water cooler and heating equipment for 3 gallons.
  - **Internal cost:** SGD 250.
  - **GST multiplier:** 1.09.
  - **Mark-up multiplier:** 1.5.
  - **Remarks / search terms:** DRINKING WATER.
