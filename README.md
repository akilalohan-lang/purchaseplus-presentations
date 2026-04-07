# PurchasePlus Presentation Generator

Generates branded PowerPoint presentations for PurchasePlus — ready to embed screen recording videos.

## Setup

```bash
pip install -r requirements.txt
```

## Generate a Presentation

```bash
python generate.py --config configs/ihg_chiang_mai.json
```

Output: `output/IHG_ChiangMai_PurchasePlus.pptx`

## Adding Videos

Each video slide has a placeholder. After generating:
1. Open the PPTX in PowerPoint
2. Go to each **[VIDEO]** slide
3. Insert > Video > This Device — select your screen recording
4. Resize to fill the content area

## Config Format

See `configs/ihg_chiang_mai.json` for a full example.
You can duplicate and edit configs for different presentations.
