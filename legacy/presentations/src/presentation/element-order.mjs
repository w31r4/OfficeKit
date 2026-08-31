export const PRESENTATION_ELEMENT_ORDER_CAPABILITY = Symbol.for("office-kit.presentation-element-order-capability");

const FILTERED_INDEX_KEYS = ["shapes", "images", "tables", "charts", "connectors", "groups", "nativeObjects"];

function ownerItems(element) {
  const items = element?.parentGroup?.children || element?.slide?.elements?.items;
  if (!Array.isArray(items)) throw new Error("Presentation element does not belong to an ordered slide or group scene stack.");
  return items;
}

function imported(element) {
  return element?.[PRESENTATION_ELEMENT_ORDER_CAPABILITY];
}

export function presentationElementOrderCapability(element) {
  const capability = imported(element);
  return capability
    ? { ...capability }
    : { sourceBound: false, known: true, editable: true, blockedReason: "" };
}

export function presentationElementStackIndex(element) {
  return ownerItems(element).indexOf(element);
}

function assertEditable(element) {
  const capability = presentationElementOrderCapability(element);
  if (capability.sourceBound && (!capability.known || !capability.editable)) {
    const detail = capability.blockedReason ? `: ${capability.blockedReason}` : ".";
    const error = new Error(`Imported presentation element z-order cannot be safely changed${detail}`);
    error.code = "unsupported_presentation_element_reorder";
    throw error;
  }
}

function sourcePrefixIsValid(items) {
  let authoredSeen = false;
  for (const item of items) {
    if (imported(item)?.sourceBound === true) {
      if (authoredSeen) return false;
    } else {
      authoredSeen = true;
    }
  }
  return true;
}

function filteredIndexUpdates(element, orderedItems) {
  const owner = element.parentGroup || element.slide;
  const updates = [];
  for (const key of FILTERED_INDEX_KEYS) {
    const items = owner?.[key]?.items;
    if (!Array.isArray(items)) continue;
    const members = new Set(items);
    const ordered = orderedItems.filter((item) => members.has(item));
    if (ordered.length !== items.length) {
      throw new Error(`Presentation ${key} index is inconsistent with its owner scene stack.`);
    }
    updates.push({ items, ordered });
  }
  return updates;
}

export function assertPresentationElementIndexes(owner, orderedItems = owner?.children || owner?.elements?.items) {
  if (!Array.isArray(orderedItems)) throw new Error("Presentation scene stack is unavailable for index validation.");
  const collections = FILTERED_INDEX_KEYS
    .map((key) => ({ key, items: owner?.[key]?.items }))
    .filter((entry) => Array.isArray(entry.items));
  const indexedItems = collections.flatMap((entry) => entry.items);
  const indexedSet = new Set(indexedItems);
  const orderedSet = new Set(orderedItems);
  const complete = indexedItems.length === orderedItems.length
    && indexedSet.size === indexedItems.length
    && orderedSet.size === orderedItems.length
    && orderedItems.every((item) => indexedSet.has(item));
  const ordered = complete && collections.every((entry) => {
    const members = new Set(entry.items);
    const expected = orderedItems.filter((item) => members.has(item));
    return expected.length === entry.items.length && expected.every((item, index) => item === entry.items[index]);
  });
  if (!complete || !ordered) {
    const grouped = Array.isArray(owner?.children);
    const error = new Error(`Presentation ${grouped ? "group" : "element"} topology changed outside the typed scene-stack lifecycle.`);
    error.code = grouped ? "presentation_group_topology_changed" : "presentation_element_topology_changed";
    throw error;
  }
  for (const item of orderedItems) {
    if (Array.isArray(item?.children)) assertPresentationElementIndexes(item, item.children);
  }
  return true;
}

function move(element, destinationIndex, placement) {
  assertEditable(element);
  const items = ownerItems(element);
  const currentIndex = items.indexOf(element);
  if (currentIndex < 0) throw new Error("Presentation element must belong to its owner before its z-order can change.");
  const next = [...items];
  next.splice(currentIndex, 1);
  const bounded = Math.max(0, Math.min(destinationIndex, next.length));
  next.splice(bounded, 0, element);
  if (!sourcePrefixIsValid(next)) {
    const error = new Error("Authored overlays on an imported slide must remain above the complete source-bound element prefix.");
    error.code = "unsupported_presentation_element_reorder";
    throw error;
  }
  const indexUpdates = filteredIndexUpdates(element, next);
  items.splice(0, items.length, ...next);
  for (const update of indexUpdates) update.items.splice(0, update.items.length, ...update.ordered);
  element._zPlacement = placement;
  return element;
}

function assertPeer(element, target) {
  if (!target || target === element) throw new TypeError("Presentation z-order target must be a different element in the same scene stack.");
  const items = ownerItems(element);
  if (ownerItems(target) !== items || !items.includes(target)) {
    throw new Error("Presentation z-order target must belong to the same slide or group scene stack.");
  }
  return items;
}

const orderingDescriptors = {
  zOrderCapability: { enumerable: true, configurable: false, get() { return presentationElementOrderCapability(this); } },
  stackIndex: { enumerable: true, configurable: false, get() { return presentationElementStackIndex(this); } },
  sendToBack: { enumerable: false, configurable: false, value() { return move(this, 0, "back"); } },
  bringToFront: { enumerable: false, configurable: false, value() { return move(this, ownerItems(this).length - 1, "front"); } },
  moveBefore: {
    enumerable: false,
    configurable: false,
    value(target) {
      const items = assertPeer(this, target);
      const targetIndex = items.indexOf(target);
      const currentIndex = items.indexOf(this);
      return move(this, targetIndex - (currentIndex < targetIndex ? 1 : 0), "custom");
    },
  },
  moveAfter: {
    enumerable: false,
    configurable: false,
    value(target) {
      const items = assertPeer(this, target);
      const targetIndex = items.indexOf(target);
      const currentIndex = items.indexOf(this);
      return move(this, targetIndex + (currentIndex > targetIndex ? 1 : 0), "custom");
    },
  },
};

export function installPresentationElementOrdering(element) {
  if (!Object.hasOwn(element, "zOrderCapability")) Object.defineProperties(element, orderingDescriptors);
  return element;
}

export function removePresentationElementFromOrder(element) {
  const items = ownerItems(element);
  const index = items.indexOf(element);
  if (index >= 0) items.splice(index, 1);
}
