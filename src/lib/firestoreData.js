import {
  collection, onSnapshot, doc, setDoc, deleteDoc, addDoc, getDocs,
} from "firebase/firestore";
import { useEffect, useState, useRef } from "react";
import { db } from "../firebase";

// Se suscribe a una colección de Firestore. La primera vez que la colección
// está vacía, la siembra con seedItems (conservando sus id originales) para
// no perder los datos de ejemplo/reales que ya veníamos usando en el prototipo.
export function useSeededCollection(name, seedItems) {
  const [items, setItems] = useState([]);
  const [ready, setReady] = useState(false);
  const seeded = useRef(false);

  useEffect(() => {
    let unsub;
    (async () => {
      if (!seeded.current) {
        seeded.current = true;
        try {
          const snap = await getDocs(collection(db, name));
          if (snap.empty && seedItems && seedItems.length) {
            await Promise.all(
              seedItems.map(({ id, ...rest }) => setDoc(doc(db, name, id), rest))
            );
          }
        } catch (err) {
          console.error(`No se pudo sembrar la colección ${name}:`, err);
        }
      }
      unsub = onSnapshot(
        collection(db, name),
        (snap) => {
          setItems(snap.docs.map((d) => ({ id: d.id, ...d.data() })));
          setReady(true);
        },
        (err) => {
          console.error(`Error escuchando ${name}:`, err);
          setReady(true);
        }
      );
    })();
    return () => unsub && unsub();
  }, [name]);

  return [items, ready];
}

// Crea o actualiza un documento. Si no se pasa id, Firestore genera uno nuevo.
export async function saveItem(name, id, data) {
  if (id) {
    await setDoc(doc(db, name, id), data);
    return id;
  }
  const ref = await addDoc(collection(db, name), data);
  return ref.id;
}

export async function deleteItem(name, id) {
  await deleteDoc(doc(db, name, id));
}

export async function replaceCollection(name, currentIds, newItems) {
  await Promise.all(currentIds.map((id) => deleteItem(name, id)));
  await Promise.all(newItems.map(({ id, ...rest }) => setDoc(doc(db, name, id), rest)));
}
