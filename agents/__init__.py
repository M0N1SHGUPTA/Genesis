"""
agents/ — Multi-agent intelligence layer for the MD-to-PPTX pipeline.

Three focused LLM agents replace the single monolithic StorylineGenerator:

  Agent 1 — ContentExtractor  : Distils the parsed markdown into key insights,
                                 visual candidates, and global statistics.
  Agent 2 — StorylinePlanner  : Decides slide count, narrative order, and layout
                                 type for each slide.
  Agent 3 — ContentTransformer: Fills every slide's content fields (bullets,
                                 chart data, table data, step lists, etc.)

The agents/pipeline.py module chains all three and applies a Python
DesignEnforcer (renderer/validator.py) before returning the final blueprint.
"""
