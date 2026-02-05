/**
 * Shared toast helper used across taskpane and commands.
 */

export function showToast(message: string, duration = 2000): void {
  let toast = document.getElementById("pi-toast");
  if (!toast) {
    toast = document.createElement("div");
    toast.id = "pi-toast";
    toast.className = "pi-toast";
    document.body.appendChild(toast);
  }
  toast.textContent = message;
  toast.classList.add("visible");
  setTimeout(() => toast!.classList.remove("visible"), duration);
}
