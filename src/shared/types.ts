export interface SelectionSnapshot {
  slideId: string;
  shapeId: string;
  start: number;
  length: number;
  text: string;
  isInsertionPoint?: boolean;
  hyperlinkAddress?: string;
  hyperlinkScreenTip?: string;
}

export interface ResourceSelection {
  url: string;
  resolvedUrl?: string;
  text?: string;
  title?: string;
}

export type BindingMode = "create-hyperlink" | "edit-hyperlink";

export interface DialogReadyMessage {
  type: "dialog-ready";
}

export interface DialogCloseMessage {
  type: "close-dialog";
}

export interface DialogApplyLinkMessage {
  type: "apply-link";
  payload: ResourceSelection;
}

export interface DialogErrorMessage {
  type: "apply-error";
  message: string;
}

export interface DialogContextMessage {
  type: "selection-context";
  selection: SelectionSnapshot;
  resourceUrl: string;
  mode: BindingMode;
}

export type ParentToDialogMessage = DialogContextMessage | DialogErrorMessage;

export type DialogToParentMessage = DialogReadyMessage | DialogCloseMessage | DialogApplyLinkMessage;
