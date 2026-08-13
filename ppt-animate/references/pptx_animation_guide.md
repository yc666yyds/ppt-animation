# PPT Animation Technical Guide

This document provides detailed technical reference for the ppt-animate skill.

## Animation Effects

### Transition Types (EntryEffect)

| Name  | ID  | Description      |
|-------|-----|------------------|
| fade  | 2   | Fade in/out      |
| wipe  | 7   | Wipe             |
| split | 13  | Split            |
| zoom  | 41  | Zoom             |
| push  | 10  | Push             |
| cover | 6   | Cover            |

### Entrance Effects (msoAnimEffect)

| Name         | ID  | Description      |
|--------------|-----|------------------|
| appear       | 1   | Instant appear   |
| checkerboard | 2   | Checkerboard     |
| circle       | 86  | Circle expand    |
| box          | 86  | Box expand       |
| spin         | 3   | Spin             |
| fly          | 4   | Fly in           |
| blend        | 17  | Blend            |
| blur         | 28  | Blur             |
| compress     | 40  | Compress         |
| dissolve     | 41  | Dissolve         |
| explode      | 42  | Explode          |
| fade         | 43  | Fade             |
| glow         | 35  | Glow             |
| grow         | 53  | Grow             |
| misty        | 59  | Misty            |
| ripple       | 64  | Ripple           |
| reveal       | 65  | Reveal           |
| roll         | 66  | Roll             |
| shrink       | 71  | Shrink           |
| swizzle      | 100 | Swizzle          |
| teeter       | 105 | Teeter           |
| typeWriter   | 106 | Type writer      |

## Trigger Types (msoAnimTrigger)

| Value | Constant                  | Description              |
|-------|---------------------------|--------------------------|
| 0     | msoAnimTriggerWithPrevious| Start with previous      |
| 1     | msoAnimTriggerAfterPrevious| Start after previous    |
| 2     | msoAnimTriggerOnClick     | Start on click           |

## COM API Reference

```python
# Add animation to a shape
seq = slide.TimeLine.MainSequence
# Clear existing animations first
while seq.Count > 0:
    seq.Item(1).Delete()

# AddEffect(shape, effectType, triggerType, triggerID)
eff = seq.AddEffect(shape, effect_type, 1, 0)  # click-triggered
eff.Timing.Duration = 2.0
```

## Correct XML Structure for Animations

```xml
<p:timing>
  <p:tnLst>
    <p:par>
      <p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot">
        <p:childTnLst>
          <p:seq concurrent="1" nextAc="seek">
            <p:cTn id="1001" dur="indefinite" nodeType="mainSeq" fill="hold">
              <!-- Required: previous condition -->
              <p:prevCondLst>
                <p:cTn condAlign="afterPp" dur="indefinite" fill="hold"/>
              </p:prevCondLst>
              <!-- Required: next condition -->
              <p:nextCondLst>
                <p:cTn condAlign="afterPp" dur="indefinite" fill="hold"/>
              </p:nextCondLst>
              <!-- Click trigger condition -->
              <p:stCondLst>
                <p:cond delay="indefinite"/>
              </p:stCondLst>
              <p:childTnLst>
                <p:par>
                  <p:cTn fill="hold">
                    <p:tgtEl>
                      <p:spTgt spid="2"/>
                    </p:tgtEl>
                    <p:animEffect transition="in" filter="appear(in)"/>
                  </p:cTn>
                </p:par>
              </p:childTnLst>
            </p:cTn>
          </p:seq>
        </p:childTnLst>
      </p:cTn>
    </p:par>
  </p:tnLst>
  <p:bldLst/>
</p:timing>
```

## Key Requirements for Valid XML Animation

1. **Each cTn must have**: `id`, `dur`, `fill`, and appropriate `nodeType`
2. **prevCondLst and nextCondLst are REQUIRED** - omitting them causes blank slides on play
3. **stCondLst** with `delay="indefinite"` creates click-triggered behavior
4. **seq must have**: `concurrent="1"` and `nextAc="seek"`
5. **spTgt spid** must match the actual shape ID in the slide XML (1-based, index 1 is usually background)

## Gencache Patch for win32com

```python
def _patch_gencache():
    """Patch win32com gencache to avoid CLSIDToClassMap errors."""
    import win32com.client.gencache as gencache_mod
    orig_add = gencache_mod.AddModuleToCache
    def patched_add(*args, **kwargs):
        try:
            return orig_add(*args, **kwargs)
        except AttributeError as e:
            if "CLSIDToClassMap" in str(e):
                return None
            raise
    gencache_mod.AddModuleToCache = patched_add
```

## Color Scheme

- Background: #0D1B2A (dark blue)
- Card fill: #1B3A5C (medium blue)
- Accent/Gold: #E8B830 (gold)
- Text white: #FFFFFF
- Text light gray: #8899AA
