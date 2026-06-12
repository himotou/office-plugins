import * as React from "react";
import { createRoot } from "react-dom/client";
import { FluentProvider, webLightTheme } from "@fluentui/react-components";
import ResourcePickerApp from "./ResourcePickerApp";

/* global document, HTMLElement */

const rootElement: HTMLElement | null = document.getElementById("container");
const root = rootElement ? createRoot(rootElement) : undefined;

root?.render(
  <FluentProvider theme={webLightTheme}>
    <ResourcePickerApp />
  </FluentProvider>
);
