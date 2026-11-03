// Función programada: pide a Netlify un nuevo despliegue (que ejecuta update.py y refresca los datos).
// Requiere la variable de entorno BUILD_HOOK_URL (Site configuration > Build & deploy > Build hooks).
// Frecuencia: cada 2 días a las 06:00 UTC. Cada despliegue de producción consume 15 créditos;
// el plan gratuito incluye 300 al mes. Para actualizar a diario, cambia a "0 6 * * *" (≈450 créditos/mes).
export default async () => {
  const url = process.env.BUILD_HOOK_URL;
  if (!url) {
    return new Response("Falta la variable BUILD_HOOK_URL", { status: 500 });
  }
  const r = await fetch(url, { method: "POST" });
  return new Response(`Despliegue solicitado: ${r.status}`);
};

export const config = { schedule: "0 6 */2 * *" };
