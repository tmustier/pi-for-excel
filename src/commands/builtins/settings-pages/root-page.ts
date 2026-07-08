/**
 * Settings root page — grouped navigation plus the handful of inline
 * behavior settings (execution mode, model-switch fork, language).
 */

import { getAppStorage } from "../../../storage/local/app-storage.js";

import { listOpenAiGatewayConfigs } from "../../../auth/custom-gateways.js";
import { PI_EXECUTION_MODE_CHANGED_EVENT, type ExecutionMode } from "../../../execution/mode.js";
import type { ModelSwitchBehavior } from "../../../models/switch-behavior.js";
import { getLanguage, initLanguage, t } from "../../../language/index.js";
import {
  Archive,
  FlaskConical,
  Keyboard,
  NotebookPen,
  Plug,
  Puzzle,
  Ruler,
  Server,
  ShieldCheck,
  Zap,
  lucide,
} from "../../../ui/lucide-icons.js";
import {
  createNavRow,
  createSettingsGroup,
  createSettingSelectRow,
  createSettingToggleRow,
} from "../../../ui/settings-rows.js";
import type { SettingsPageContext, SettingsShellPage } from "../../../ui/settings-shell.js";
import { showToast } from "../../../ui/toast.js";
import { getSettingsPagesDependencies } from "./dependencies.js";

function isRootPagePayloadShape(value: DynamicValue): value is DynamicObject {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function buildProvidersGroup(ctx: SettingsPageContext): HTMLElement {
  const group = createSettingsGroup(t("settings.group.ai_providers"));

  const providersRow = createNavRow({
    icon: lucide(Zap),
    label: t("settings.row.providers"),
    sublabel: t("settings.row.providers.sub"),
    onActivate: () => ctx.navigate("providers"),
  });

  const gatewayRow = createNavRow({
    icon: lucide(Server),
    label: t("settings.row.gateway"),
    sublabel: t("settings.row.gateway.sub"),
    onActivate: () => ctx.navigate("gateway"),
  });

  const proxyRow = createNavRow({
    icon: lucide(ShieldCheck),
    label: t("settings.row.proxy"),
    sublabel: t("settings.row.proxy.sub"),
    onActivate: () => ctx.navigate("proxy"),
  });

  group.list.append(providersRow.root, gatewayRow.root, proxyRow.root);

  // Best-effort async value previews; never block rendering.
  void (async () => {
    const storage = getAppStorage();

    try {
      const configured = await storage.providerKeys.list();
      providersRow.setValue(
        configured.length > 0
          ? t("settings.value.connected_count", { count: configured.length })
          : t("settings.value.not_set_up"),
      );
    } catch {
      // leave preview empty
    }

    try {
      const gateways = await listOpenAiGatewayConfigs(storage.customProviders);
      gatewayRow.setValue(
        gateways.length > 0 ? String(gateways.length) : t("settings.value.none"),
      );
    } catch {
      // leave preview empty
    }

    try {
      const proxyEnabled = await storage.settings.get<boolean>("proxy.enabled");
      proxyRow.setValue(proxyEnabled === true ? t("settings.value.on") : t("settings.value.off"));
    } catch {
      // leave preview empty
    }
  })();

  return group.root;
}

function buildBehaviorGroup(ctx: SettingsPageContext): HTMLElement {
  const deps = getSettingsPagesDependencies();
  const group = createSettingsGroup(t("settings.group.behavior"));

  // ── Auto-apply (execution mode) ──
  const autoApply = createSettingToggleRow({
    label: t("settings.section.execution.auto_mode"),
    sublabel: t("settings.section.execution.auto_sublabel"),
  });

  const getExecutionMode = deps.getExecutionMode;
  const setExecutionMode = deps.setExecutionMode;

  if (!getExecutionMode || !setExecutionMode) {
    autoApply.input.disabled = true;
  } else {
    let currentMode = getExecutionMode();
    autoApply.input.checked = currentMode === "yolo";

    autoApply.input.addEventListener("change", () => {
      const nextMode: ExecutionMode = autoApply.input.checked ? "yolo" : "safe";
      if (nextMode === currentMode) return;

      autoApply.input.disabled = true;
      void setExecutionMode(nextMode).then(
        () => {
          currentMode = nextMode;
          showToast(nextMode === "yolo" ? t("settings.toast.auto_mode") : t("settings.toast.confirm_mode"));
        },
        () => {
          autoApply.input.checked = currentMode === "yolo";
          showToast(t("settings.toast.execution_failed"));
        },
      ).finally(() => {
        autoApply.input.disabled = false;
      });
    });

    const onExecutionModeChanged = (event: Event): void => {
      if (!(event instanceof CustomEvent)) return;
      const detail: DynamicValue = event.detail;
      if (!isRootPagePayloadShape(detail)) return;
      const mode = detail.mode;
      if (mode !== "yolo" && mode !== "safe") return;
      currentMode = mode;
      autoApply.input.checked = mode === "yolo";
    };

    document.addEventListener(PI_EXECUTION_MODE_CHANGED_EVENT, onExecutionModeChanged);
    ctx.addCleanup(() => {
      document.removeEventListener(PI_EXECUTION_MODE_CHANGED_EVENT, onExecutionModeChanged);
    });
  }

  // ── Fork model switch ──
  const forkToggle = createSettingToggleRow({
    label: t("settings.section.advanced.fork_label"),
    sublabel: t("settings.section.advanced.fork_sublabel"),
  });

  const getModelSwitchBehavior = deps.getModelSwitchBehavior;
  const setModelSwitchBehavior = deps.setModelSwitchBehavior;

  if (!getModelSwitchBehavior || !setModelSwitchBehavior) {
    forkToggle.input.disabled = true;
  } else {
    let currentBehavior = getModelSwitchBehavior();
    forkToggle.input.checked = currentBehavior === "fork";

    forkToggle.input.addEventListener("change", () => {
      const nextBehavior: ModelSwitchBehavior = forkToggle.input.checked ? "fork" : "inPlace";
      if (nextBehavior === currentBehavior) return;

      forkToggle.input.disabled = true;
      void setModelSwitchBehavior(nextBehavior).then(
        () => {
          currentBehavior = nextBehavior;
          showToast(nextBehavior === "fork" ? t("settings.toast.fork_on") : t("settings.toast.fork_off"));
        },
        () => {
          forkToggle.input.checked = currentBehavior === "fork";
          showToast(t("settings.toast.fork_failed"));
        },
      ).finally(() => {
        forkToggle.input.disabled = false;
      });
    });
  }

  // ── Language ──
  const languageRow = createSettingSelectRow({
    label: t("settings.section.language.label"),
    options: [
      { value: "en", label: t("settings.section.language.en") },
      // Disclose that the Chinese translation is AI-generated (issue #608).
      { value: "zh-CN", label: t("settings.section.language.zh") },
    ],
    value: getLanguage(),
    onChange: (newLang) => {
      initLanguage(newLang);
      void (async () => {
        try {
          const storage = getAppStorage();
          await storage.settings.set("language", newLang);
          showToast(t("settings.lang.reloading"));
          setTimeout(() => location.reload(), 1000);
        } catch {
          showToast(t("settings.lang.saveFailed"));
        }
      })();
    },
  });

  group.list.append(autoApply.root, forkToggle.root, languageRow.root);
  return group.root;
}

function buildRulesDataGroup(ctx: SettingsPageContext): HTMLElement {
  const group = createSettingsGroup(t("settings.group.rules_data"));

  const rulesRow = createNavRow({
    icon: lucide(Ruler),
    label: t("settings.row.rules"),
    sublabel: t("settings.row.rules.sub"),
    onActivate: () => ctx.navigate("rules"),
  });

  const backupsRow = createNavRow({
    icon: lucide(Archive),
    label: t("settings.row.backups"),
    sublabel: t("settings.row.backups.sub"),
    onActivate: () => ctx.navigate("backups"),
  });

  group.list.append(rulesRow.root, backupsRow.root);
  return group.root;
}

function buildExtensionsGroup(ctx: SettingsPageContext): HTMLElement {
  const deps = getSettingsPagesDependencies();
  const group = createSettingsGroup(t("settings.group.extensions"));

  const connectionsRow = createNavRow({
    icon: lucide(Plug),
    label: t("settings.row.connections"),
    sublabel: t("settings.row.connections.sub"),
    onActivate: () => ctx.navigate("connections"),
  });

  const pluginsRow = createNavRow({
    icon: lucide(Puzzle),
    label: t("settings.row.plugins"),
    sublabel: t("settings.row.plugins.sub"),
    onActivate: () => ctx.navigate("plugins"),
  });

  const skillsRow = createNavRow({
    icon: lucide(NotebookPen),
    label: t("settings.row.skills"),
    sublabel: t("settings.row.skills.sub"),
    onActivate: () => ctx.navigate("skills"),
  });

  const extensionManager = deps.extensionsHub?.extensionManager;
  if (extensionManager) {
    try {
      const installed = extensionManager.list().length;
      if (installed > 0) {
        pluginsRow.setValue(t("settings.value.installed_count", { count: installed }));
      }
    } catch {
      // leave preview empty
    }
  }

  group.list.append(connectionsRow.root, pluginsRow.root, skillsRow.root);
  return group.root;
}

function buildHelpGroup(ctx: SettingsPageContext): HTMLElement {
  const group = createSettingsGroup(t("settings.group.help"));

  const shortcutsRow = createNavRow({
    icon: lucide(Keyboard),
    label: t("settings.row.shortcuts"),
    onActivate: () => ctx.navigate("shortcuts"),
  });

  const experimentalRow = createNavRow({
    icon: lucide(FlaskConical),
    label: t("settings.row.experimental"),
    sublabel: t("settings.row.experimental.sub"),
    onActivate: () => ctx.navigate("experimental"),
  });

  group.list.append(shortcutsRow.root, experimentalRow.root);
  return group.root;
}

export function createRootPage(): SettingsShellPage {
  return {
    id: "root",
    title: () => t("settings.title"),
    render: (ctx) => {
      ctx.body.append(
        buildProvidersGroup(ctx),
        buildBehaviorGroup(ctx),
        buildRulesDataGroup(ctx),
        buildExtensionsGroup(ctx),
        buildHelpGroup(ctx),
      );
    },
  };
}
