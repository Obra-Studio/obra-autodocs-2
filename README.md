# Obra Autodocs - Figma Plugin

Made by [Obra Studio](https://obra.studio/).

A Figma plugin that analyzes a component variant set and documents the different properties.

## How it works

The plugin looks at the different properties in a component variant set, and renders a grid of those properties with visible labels on the sides.

For example a component has an enumerated property called Type, and it has 3 values: Type 1, Type 2 and Type 3. It has another boolean property called Active which has 2 values: On and Off (other ways to denote booleans in Figma are True and False).

The plugin will first analyze the ideal pattern for laying out the grid. It will analyze the size of each item in the grid and determine an ideal grid.

The default grid style is based on the default style of Figma for component variant sets, which is using the color #9747FF as a dashed border.

The plugin has options to show booleans and show nested properties. For the first iteration, when there are nested properties, ignore them. Don't show booleans either; focus on showing enumerated properties.

When multiple properties are shown on top of each other, you can use brackets to group properties together.

## Installation

1. In Figma, go to **Plugins** > **Development** > **Import plugin from manifest...**
2. Navigate to this folder and select the `manifest.json` file
3. The plugin will appear in your Plugins > Development menu

## Technical Details

- Written in ES5 and vanilla JavaScript
- Uses Figma's Plugin API

## File Structure

```
ziptility-icon-treatment/
├── manifest.json    # Plugin configuration
├── code.js          # Main plugin logic (runs in Figma sandbox)
├── ui.html          # User interface (runs in iframe)
└── README.md        # This file
```

## License

MIT
