// Deck layer: theme + layout types.
//
// A theme is 12 symbolic color slots + a major/minor font pair. A layout is a
// named set of placeholder definitions (position + default styling) that
// `addSlide({ layoutId })` materializes into real elements.

import type {
  Frame,
  Color,
  Fill,
  ThemeColorSlot,
  PlaceholderType,
  BodyStyle,
  SlideElement,
} from '../scene/types';
import type { Bullet, TextAlign } from '../text/model';

export interface ThemeData {
  name: string;
  /** Slot → literal hex, e.g. '#1A1A1A'. */
  colors: Record<ThemeColorSlot, string>;
  fonts: {
    /** Headings font family. */
    major: string;
    /** Body font family. */
    minor: string;
  };
}

/** Default styling applied when a placeholder is materialized into an element. */
export interface PlaceholderStyle {
  color?: Color;
  fontSize?: number;
  bold?: boolean;
  align?: TextAlign;
  vAlign?: BodyStyle['vAlign'];
  /** 'major' | 'minor' | a literal family. */
  font?: string;
}

export interface PlaceholderDef {
  type: PlaceholderType;
  /** Distinguishes multiple placeholders of the same type on a layout. */
  idx: number;
  frame: Frame;
  /** Prompt text shown when the materialized element is empty. */
  prompt?: string;
  style?: PlaceholderStyle;
}

/**
 * The text defaults of one paragraph level (0 to 8) of a placeholder type, flattened from the master's `p:titleStyle` /
 * `p:bodyStyle` and the layout's `a:lstStyle` (PPTX). Used when a placeholder is materialized and kept so an importer can
 * hand the effective values to slide text.
 */
export interface LevelTextStyle extends PlaceholderStyle {
  italic?: boolean;
  bullet?: Bullet | null;
  /** Left margin of the level in px (`marL`). */
  marginLeft?: number;
  /** First-line indent in px (`indent`, negative for a hanging bullet). */
  firstLineIndent?: number;
  /** Line height multiplier. */
  lineHeight?: number;
  /** Space before in px. */
  spaceBefore?: number;
}

/** Per placeholder type, the level styles; index = paragraph level, a missing level inherits the nearest lower one. */
export type LayoutTextStyles = Partial<Record<PlaceholderType, LevelTextStyle[]>>;

export interface LayoutData {
  id: string;
  name: string;
  /** Name of the master the layout came from (PPTX `p:cSld/@name` of the master), for display and re-export. */
  masterName?: string;
  background?: Fill;
  placeholders: PlaceholderDef[];
  /** Static, non-placeholder content from the layout and its master (logos, rules, footers), drawn under slide content. */
  elements?: SlideElement[];
  textStyles?: LayoutTextStyles;
}
