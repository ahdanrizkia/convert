import {
  CATEGORIES,
  FORMAT_CATEGORY_MAP,
  FORMAT_ROUTES,
  MIME_TYPES,
  normalizeFormat,
  type CategoryId,
} from './definitions';
import { convertImage } from './imageConverter';
import { convertData } from './dataConverter';
import { convertDocument } from './documentConverter';
import { convertMedia } from './mediaConverter';

export type ConverterHandler = (
  inputPath: string,
  outputPath: string,
  sourceFormat: string,
  targetFormat: string
) => Promise<void>;

export interface RegisteredRoute {
  key: string;
  category: CategoryId;
  phase: 1 | 2 | 3;
  active: boolean;
  sourceFormat: string;
  targetFormat: string;
  targetMimeType: string;
  handler?: ConverterHandler;
}

const routeRegistry = new Map<string, RegisteredRoute>();

function makeRouteKey(sourceFormat: string, targetFormat: string): string {
  return `${normalizeFormat(sourceFormat)}->${normalizeFormat(targetFormat)}`;
}

/**
 * Registers a conversion route in the modular converter registry.
 */
export function registerConverterRoute(route: Omit<RegisteredRoute, 'key'>): void {
  const key = makeRouteKey(route.sourceFormat, route.targetFormat);
  routeRegistry.set(key, {
    ...route,
    sourceFormat: normalizeFormat(route.sourceFormat),
    targetFormat: normalizeFormat(route.targetFormat),
    key,
  });
}

// Bootstrap all active routes with their specialized converters
for (const [source, targets] of Object.entries(FORMAT_ROUTES)) {
  const category = FORMAT_CATEGORY_MAP[source];
  if (!category) continue;

  for (const target of targets) {
    let handler: ConverterHandler;
    let phase: 1 | 2 | 3 = 1;

    if (category === 'media' || target === 'mp3') {
      handler = convertMedia;
      phase = 3;
    } else if (
      category === 'documents' ||
      target === 'pdf' ||
      target === 'docx' ||
      target === 'pptx' ||
      target === 'xlsx' ||
      source === 'docx' ||
      source === 'pdf' ||
      source === 'xlsx' ||
      source === 'pptx'
    ) {
      handler = convertDocument;
      phase = 2;
    } else if (category === 'images') {
      handler = convertImage;
      phase = 1;
    } else {
      handler = convertData;
      phase = 1;
    }

    registerConverterRoute({
      category,
      phase,
      active: true,
      sourceFormat: source,
      targetFormat: target,
      targetMimeType: MIME_TYPES[target] ?? 'application/octet-stream',
      handler,
    });
  }
}

// Pre-register remaining inactive category slots as placeholders
for (const cat of CATEGORIES) {
  if (cat.active || !cat.plannedRoutes) continue;
  for (const spec of cat.plannedRoutes) {
    const [rawSrc, rawDst] = spec.split('->').map((s) => s.trim().toLowerCase());
    if (rawSrc && rawDst) {
      registerConverterRoute({
        category: cat.id,
        phase: cat.phase,
        active: false,
        sourceFormat: rawSrc,
        targetFormat: rawDst,
        targetMimeType: 'application/octet-stream',
      });
    }
  }
}

/**
 * Resolves a converter route for a given `(sourceFormat, targetFormat)` pair.
 */
export function resolveConverterRoute(
  sourceFormat: string,
  targetFormat: string
): RegisteredRoute | null {
  return routeRegistry.get(makeRouteKey(sourceFormat, targetFormat)) ?? null;
}

/**
 * Lists all registered routes (useful for diagnostics or future capabilities discovery).
 */
export function listRegisteredRoutes(): RegisteredRoute[] {
  return Array.from(routeRegistry.values());
}
