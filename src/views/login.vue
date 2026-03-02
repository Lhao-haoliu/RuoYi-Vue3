<template>
  <div class="portal-login">
    <div class="login-panel">
      <div class="login-brand">
        <img :src="portalLogo" alt="SEMI Logo" />
        <h2>{{ title }}</h2>
        <p>SEMI控制台</p>
      </div>
      <el-form ref="loginRef" :model="loginForm" :rules="loginRules" class="login-form">
        <el-form-item prop="username">
          <el-input v-model="loginForm.username" size="large" placeholder="请输入账号">
            <template #prefix><svg-icon icon-class="user" class="input-icon" /></template>
          </el-input>
        </el-form-item>
        <el-form-item prop="password">
          <el-input v-model="loginForm.password" type="password" size="large" placeholder="请输入密码" @keyup.enter="handleLogin">
            <template #prefix><svg-icon icon-class="password" class="input-icon" /></template>
          </el-input>
        </el-form-item>
        <el-form-item prop="code" v-if="captchaEnabled">
          <el-input v-model="loginForm.code" size="large" placeholder="请输入验证码" class="code-input" @keyup.enter="handleLogin">
            <template #prefix><svg-icon icon-class="validCode" class="input-icon" /></template>
          </el-input>
          <img :src="codeUrl" @click="getCode" class="login-code-img" />
        </el-form-item>
        <div class="login-row">
          <el-checkbox v-model="loginForm.rememberMe">记住密码</el-checkbox>
          <router-link v-if="register" class="register-link" :to="'/register'">立即注册</router-link>
        </div>
        <el-button :loading="loading" size="large" class="portal-primary-btn submit-btn" @click.prevent="handleLogin">
          {{ loading ? '登录中...' : '登录' }}
        </el-button>
      </el-form>
    </div>
  </div>
</template>

<script setup>
import { getCodeImg } from "@/api/login"
import Cookies from "js-cookie"
import { encrypt, decrypt } from "@/utils/jsencrypt"
import useUserStore from '@/store/modules/user'
import portalLogo from '@/assets/logo/company-logo.svg'

const title = import.meta.env.VITE_APP_TITLE
const userStore = useUserStore()
const route = useRoute()
const router = useRouter()
const { proxy } = getCurrentInstance()

const loginForm = ref({
  username: "admin",
  password: "admin123",
  rememberMe: false,
  code: "",
  uuid: ""
})

const loginRules = {
  username: [{ required: true, trigger: "blur", message: "请输入您的账号" }],
  password: [{ required: true, trigger: "blur", message: "请输入您的密码" }],
  code: [{ required: true, trigger: "change", message: "请输入验证码" }]
}

const codeUrl = ref("")
const loading = ref(false)
const captchaEnabled = ref(true)
const register = ref(false)
const redirect = ref(undefined)

watch(route, (newRoute) => {
  redirect.value = newRoute.query && newRoute.query.redirect
}, { immediate: true })

function handleLogin() {
  proxy.$refs.loginRef.validate(valid => {
    if (!valid) return
    loading.value = true
    if (loginForm.value.rememberMe) {
      Cookies.set("username", loginForm.value.username, { expires: 30 })
      Cookies.set("password", encrypt(loginForm.value.password), { expires: 30 })
      Cookies.set("rememberMe", loginForm.value.rememberMe, { expires: 30 })
    } else {
      Cookies.remove("username")
      Cookies.remove("password")
      Cookies.remove("rememberMe")
    }
    userStore.login(loginForm.value).then(() => {
      const query = route.query
      const otherQueryParams = Object.keys(query).reduce((acc, cur) => {
        if (cur !== "redirect") acc[cur] = query[cur]
        return acc
      }, {})
      router.push({ path: redirect.value || "/", query: otherQueryParams })
    }).catch(() => {
      loading.value = false
      if (captchaEnabled.value) getCode()
    })
  })
}

function getCode() {
  getCodeImg().then(res => {
    captchaEnabled.value = res.captchaEnabled === undefined ? true : res.captchaEnabled
    if (captchaEnabled.value) {
      codeUrl.value = "data:image/gif;base64," + res.img
      loginForm.value.uuid = res.uuid
    }
  })
}

function getCookie() {
  const username = Cookies.get("username")
  const password = Cookies.get("password")
  const rememberMe = Cookies.get("rememberMe")
  loginForm.value = {
    username: username === undefined ? loginForm.value.username : username,
    password: password === undefined ? loginForm.value.password : decrypt(password),
    rememberMe: rememberMe === undefined ? false : Boolean(rememberMe)
  }
}

getCode()
getCookie()
</script>

<style scoped lang="scss">
.portal-login {
  min-height: 100vh;
  display: grid;
  place-items: center;
  padding: 24px 16px;
  background-color: #ffffff;
  background-image:
    linear-gradient(to right, rgba(156, 163, 175, 0.15) 1px, transparent 1px),
    linear-gradient(to bottom, rgba(156, 163, 175, 0.15) 1px, transparent 1px);
  background-size: 20px 20px;
}

.login-panel {
  width: min(440px, 100%);
  padding: 30px 28px;
  border-radius: 6px;
  background: #ffffff;
  border: 1px solid #e5e7eb;
  box-shadow: none;
}

.login-brand {
  text-align: center;
  margin-bottom: 20px;
  img {
    width: 120px;
    height: 36px;
    object-fit: contain;
  }
  h2 {
    margin: 10px 0 6px;
    font-size: 22px;
    color: #111827;
  }
  p {
    margin: 0;
    color: #9ca3af;
    font-size: 13px;
  }
}

.input-icon {
  width: 14px;
  height: 14px;
}

.code-input {
  width: calc(100% - 130px);
}

.login-code-img {
  width: 116px;
  height: 42px;
  margin-left: 12px;
  border-radius: 6px;
  cursor: pointer;
  border: 1px solid #e5e7eb;
}

.login-row {
  display: flex;
  justify-content: space-between;
  align-items: center;
  margin: 2px 0 18px;
}

.register-link {
  color: #2563eb;
  font-size: 13px;
}

.submit-btn {
  width: 100%;
}

.login-form {
  :deep(.el-form-item) {
    margin-bottom: 16px;
  }

  :deep(.el-input__wrapper) {
    min-height: 42px;
    border-radius: 6px;
    border: 1px solid #e5e7eb;
    background: #f3f4f6;
    box-shadow: none;
  }

  :deep(.el-input__wrapper.is-focus) {
    border-color: #2563eb;
    box-shadow: 0 0 0 1px rgb(37 99 235 / 0.2);
  }

  :deep(.el-input__inner) {
    color: #111827;
  }
}
</style>

