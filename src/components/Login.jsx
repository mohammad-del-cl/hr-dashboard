import { useState } from "react"

import {
  LayoutDashboard,
  User,
  Lock,
  Eye,
  EyeOff,
  LogIn,
  AlertCircle,
} from "lucide-react"

import { USERNAME, PASSWORD } from "../config/authConfig"

/* =========================================================
   Login
   صفحه ورود - فقط سمت Frontend (بدون Backend)
========================================================= */

function Login({ onLoginSuccess }) {
  const [username, setUsername] = useState("")
  const [password, setPassword] = useState("")
  const [showPassword, setShowPassword] = useState(false)
  const [error, setError] = useState("")
  const [isSubmitting, setIsSubmitting] = useState(false)

  function handleSubmit(event) {
    event.preventDefault()

    if (isSubmitting) {
      return
    }

    const trimmedUsername = username.trim()
    const trimmedPassword = password.trim()

    if (!trimmedUsername || !trimmedPassword) {
      setError("لطفاً نام کاربری و رمز عبور را وارد کنید.")
      return
    }

    setIsSubmitting(true)
    setError("")

    // یک تأخیر کوچک فقط برای حس بهتر (UX)؛ منطق واقعی همان مقایسه ساده زیر است.
    setTimeout(() => {
      if (
        trimmedUsername === USERNAME &&
        trimmedPassword === PASSWORD
      ) {
        onLoginSuccess()
        return
      }

      setError("نام کاربری یا رمز عبور اشتباه است.")
      setIsSubmitting(false)
    }, 250)
  }

  return (
    <div
      dir="rtl"
      className="flex min-h-screen items-center justify-center bg-[#1a1a1a] px-4 text-white"
    >
      <div className="w-full max-w-sm rounded-2xl border border-[#d4a017]/30 bg-[#151515] p-8 shadow-2xl">
        {/* =================================================
            Logo / Title
        ================================================= */}

        <div className="mb-6 flex flex-col items-center text-center">
          <div className="mb-4 flex h-16 w-16 items-center justify-center rounded-2xl bg-[#d4a017] text-black">
            <LayoutDashboard size={30} />
          </div>

          <h1 className="text-xl font-black text-white">
            داشبورد منابع انسانی
          </h1>

          <p className="mt-1 text-xs text-[#d4a017]">
            برای ادامه، وارد حساب کاربری خود شوید
          </p>
        </div>

        {/* =================================================
            Form
        ================================================= */}

        <form onSubmit={handleSubmit} className="space-y-4" noValidate>
          <div>
            <label className="mb-1.5 block text-xs font-medium text-gray-400">
              نام کاربری
            </label>

            <div className="relative">
              <User
                size={18}
                className="pointer-events-none absolute right-3 top-1/2 -translate-y-1/2 text-gray-500"
              />

              <input
                type="text"
                value={username}
                onChange={(event) => {
                  setUsername(event.target.value)
                  if (error) setError("")
                }}
                autoFocus
                autoComplete="username"
                placeholder="نام کاربری"
                className="w-full rounded-xl border border-white/10 bg-[#202020] py-2.5 pr-10 pl-4 text-sm outline-none transition focus:border-[#d4a017]"
              />
            </div>
          </div>

          <div>
            <label className="mb-1.5 block text-xs font-medium text-gray-400">
              رمز عبور
            </label>

            <div className="relative">
              <Lock
                size={18}
                className="pointer-events-none absolute right-3 top-1/2 -translate-y-1/2 text-gray-500"
              />

              <input
                type={showPassword ? "text" : "password"}
                value={password}
                onChange={(event) => {
                  setPassword(event.target.value)
                  if (error) setError("")
                }}
                autoComplete="current-password"
                placeholder="رمز عبور"
                className="w-full rounded-xl border border-white/10 bg-[#202020] py-2.5 pr-10 pl-10 text-sm outline-none transition focus:border-[#d4a017]"
              />

              <button
                type="button"
                onClick={() => setShowPassword((prev) => !prev)}
                tabIndex={-1}
                className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-500 transition hover:text-[#f0c040]"
              >
                {showPassword ? (
                  <EyeOff size={18} />
                ) : (
                  <Eye size={18} />
                )}
              </button>
            </div>
          </div>

          {error && (
            <div className="flex items-center gap-2 rounded-lg border border-red-500/30 bg-red-500/10 px-3 py-2 text-xs text-red-400">
              <AlertCircle size={16} className="shrink-0" />
              <span>{error}</span>
            </div>
          )}

          <button
            type="submit"
            disabled={isSubmitting}
            className="flex w-full items-center justify-center gap-2 rounded-xl bg-[#d4a017] px-4 py-3 text-sm font-bold text-black transition hover:bg-[#f0c040] disabled:cursor-not-allowed disabled:opacity-60"
          >
            <LogIn size={18} />
            {isSubmitting ? "در حال بررسی..." : "ورود"}
          </button>
        </form>

        <p className="mt-6 text-center text-[11px] leading-5 text-gray-600">
          این ورود صرفاً یک محافظت ساده سمت مرورگر است و جایگزین
          یک سیستم احراز هویت واقعی (Backend) نیست.
        </p>
      </div>
    </div>
  )
}

export default Login
