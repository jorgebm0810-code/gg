// Función programada: pide a Netlify un nuevo despliegue (que ejecuta update.py y refresca los datos).
// Requiere la variable de entorno BUILD_HOOK_URL (Site configuration > Build & deploy > Build hooks).
export default async () => {
  const url = process.env.BUILD_HOOK_URL;
  if (!url) {
    return new Response("Falta la variable BUILD_HOOK_URL", { status: 500 });
  }
  const r = await fetch(url, { method: "POST" });
  return new Response(`Despliegue solicitado: ${r.status}`);
};

export const config = { schedule: "0 6 */2 * *" };
