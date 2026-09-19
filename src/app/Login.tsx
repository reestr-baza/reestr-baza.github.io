import { useState } from 'react';
import { login } from '../storage/backend';
import { BrandMark } from '../ui/icons';

/** Вход в рабочую версию. В демо (без сервера) не показывается. */
export function Login({ onDone }: { onDone: () => void }) {
  const [user, setUser] = useState('');
  const [password, setPassword] = useState('');
  const [error, setError] = useState('');
  const [busy, setBusy] = useState(false);

  const submit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!user.trim() || !password) {
      setError('Введите логин и пароль');
      return;
    }
    setBusy(true);
    setError('');
    try {
      await login(user.trim(), password);
      onDone();
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Не удалось войти');
      setBusy(false);
    }
  };

  return (
    <main className="login">
      <form className="login-card" onSubmit={submit} noValidate>
        <div className="login-brand">
          <BrandMark size={28} />
          <span className="brand-name">Реестр</span>
        </div>
        <h1 className="login-title">Вход в базу</h1>
        <label className="form-row">
          <span className="form-label">Логин</span>
          <input
            className="input"
            name="username"
            autoComplete="username"
            autoFocus
            value={user}
            onChange={(e) => setUser(e.target.value)}
            aria-invalid={!!error}
          />
        </label>
        <label className="form-row">
          <span className="form-label">Пароль</span>
          <input
            className="input"
            type="password"
            name="password"
            autoComplete="current-password"
            value={password}
            onChange={(e) => setPassword(e.target.value)}
            aria-invalid={!!error}
            aria-describedby={error ? 'login-err' : undefined}
          />
        </label>
        {error && (
          <p id="login-err" className="form-error" role="alert">
            {error}
          </p>
        )}
        <button type="submit" className="btn btn--primary login-submit" disabled={busy}>
          {busy ? 'Входим…' : 'Войти'}
        </button>
      </form>
    </main>
  );
}
