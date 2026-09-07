// Carga el script de Google Maps una sola vez y reutiliza la misma promesa
// para todos los componentes que lo necesiten.
let mapsPromise = null;

export function loadGoogleMaps() {
  if (mapsPromise) return mapsPromise;

  mapsPromise = new Promise((resolve, reject) => {
    if (window.google && window.google.maps) {
      resolve(window.google.maps);
      return;
    }
    const apiKey = import.meta.env.VITE_GOOGLE_MAPS_API_KEY;
    if (!apiKey) {
      reject(new Error("Falta VITE_GOOGLE_MAPS_API_KEY en el archivo .env"));
      return;
    }
    const callbackName = "__onGoogleMapsLoaded";
    window[callbackName] = () => {
      resolve(window.google.maps);
      delete window[callbackName];
    };
    const script = document.createElement("script");
    script.src = `https://maps.googleapis.com/maps/api/js?key=${apiKey}&callback=${callbackName}&loading=async`;
    script.async = true;
    script.onerror = () => reject(new Error("No se pudo cargar el script de Google Maps"));
    document.head.appendChild(script);
  });

  return mapsPromise;
}

// Convierte una dirección de texto en coordenadas usando el servicio de Geocoding.
export async function geocodeAddress(address) {
  const maps = await loadGoogleMaps();
  return new Promise((resolve) => {
    const geocoder = new maps.Geocoder();
    geocoder.geocode({ address }, (results, status) => {
      if (status === "OK" && results[0]) {
        const loc = results[0].geometry.location;
        resolve({ lat: loc.lat(), lng: loc.lng() });
      } else {
        resolve(null);
      }
    });
  });
}
