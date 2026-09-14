import { Link } from "react-router-dom";

export default function NotFound() {
  return (
    <main className="flex min-h-screen items-center justify-center bg-zinc-50 px-4">
      <div className="text-center">
        <p className="text-sm font-medium text-zinc-500">404</p>

        <h1 className="mt-2 text-3xl font-semibold text-zinc-900">
          Page not found
        </h1>

        <p className="mt-3 text-zinc-600">
          The page you are looking for does not exist.
        </p>

        <Link
          to="/"
          className="mt-6 inline-flex h-10 items-center justify-center rounded-md bg-zinc-900 px-4 text-sm font-medium text-white transition hover:bg-zinc-800"
        >
          Back to home
        </Link>
      </div>
    </main>
  );
}
