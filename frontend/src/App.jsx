import { Navigate, Route, Routes } from "react-router-dom";

import Home from "./pages/Home";
import TaxEntry from "./pages/TaxEntry";
import NotFound from "./pages/NotFound";

export default function App() {
  return (
    <Routes>
      <Route path="/" element={<Home />} />

      <Route path="/tax-entry" element={<TaxEntry />} />

      <Route path="/home" element={<Navigate to="/" replace />} />

      <Route path="*" element={<NotFound />} />
    </Routes>
  );
}
