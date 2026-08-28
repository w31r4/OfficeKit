export const PRESENTATION_ELEMENT_ORDER_CAPABILITY = Symbol.for("office-kit.presentation-element-order-capability");

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
  items.splice(0, items.length, ...next);
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
