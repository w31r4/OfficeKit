const PRESENTATION_STATE = Symbol.for("office-kit.presentation-state");

export const PRESENTATION_ELEMENT_DELETION_CAPABILITY = Symbol.for("office-kit.presentation-element-deletion-capability");
export const PRESENTATION_ELEMENT_DELETED = Symbol.for("office-kit.presentation-element-deleted");

export function presentationElementDeletionCapability(element, kind) {
  const imported = element[PRESENTATION_ELEMENT_DELETION_CAPABILITY];
  if (imported) return { ...imported };
  if (element.slide?.presentation?.[PRESENTATION_STATE]) {
    return {
      sourceBound: true,
      known: false,
      supported: false,
      blockedReason: `This imported ${kind} is outside the bounded top-level element deletion profile.`,
      nativeId: undefined,
    };
  }
  return { sourceBound: false, known: true, supported: true, blockedReason: "", nativeId: undefined };
}

export function deletePresentationElement(element, collection, kind) {
  const index = collection?.items?.indexOf(element) ?? -1;
  if (index < 0) throw new Error(`Presentation ${kind} must belong to its slide before it can be deleted.`);

  const owner = element.parentGroup;
  const connectors = owner?.connectors?.items || element.slide?.connectors?.items || [];
  if (connectors.some((connector) => connector.startTargetId === element.id || connector.endTargetId === element.id)) {
    const error = new Error(`Presentation ${kind} ${element.id} cannot be deleted while a connector targets it.`);
    error.code = "unsupported_presentation_element_delete";
    throw error;
  }
  if ((element.slide?.comments?.items || []).some((thread) => thread.targetId === element.id)) {
    const error = new Error(`Presentation ${kind} ${element.id} cannot be deleted while a comment targets it.`);
    error.code = "unsupported_presentation_element_delete";
    throw error;
  }

  const capability = element.deletionCapability;
  if (capability.sourceBound && (!capability.known || !capability.supported)) {
    const detail = capability.blockedReason ? `: ${capability.blockedReason}` : ".";
    const error = new Error(`Imported presentation ${kind} cannot be safely deleted${detail}`);
    error.code = "unsupported_presentation_element_delete";
    throw error;
  }
  if (capability.sourceBound) Object.defineProperty(element, PRESENTATION_ELEMENT_DELETED, { value: true });
  collection.items.splice(index, 1);
  if (owner) {
    const childIndex = owner.children.indexOf(element);
    if (childIndex >= 0) owner.children.splice(childIndex, 1);
  }
  return element;
}
