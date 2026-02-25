<template>
  <div class="portal-login">
    <div class="login-panel">
      <div class="login-brand">
        <img :src="portalLogo" alt="Portal" />
        <h2>{{ title }}</h2>
        <p>Enterprise Intelligence Platform</p>
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
import portalLogo from '@/assets/logo/portal-logo.svg'

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
  padding: 24px;
  background:
    radial-gradient(circle at 12% 15%, rgba(74, 144, 226, 0.34), transparent 38%),
    radial-gradient(circle at 82% 20%, rgba(108, 93, 211, 0.28), transparent 32%),
    linear-gradient(160deg, #0f1f3d 0%, #132d57 55%, #1b3f77 100%);
}

.login-panel {
  width: min(440px, 100%);
  padding: 30px 28px;
  border-radius: 18px;
  background: rgba(255, 255, 255, 0.96);
  box-shadow: 0 22px 44px rgba(8, 22, 48, 0.35);
}

.login-brand {
  text-align: center;
  margin-bottom: 20px;
  img {
    width: 46px;
    height: 46px;
  }
  h2 {
    margin: 10px 0 6px;
    font-size: 22px;
    color: #0f1f3d;
  }
  p {
    margin: 0;
    color: #657291;
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
  height: 40px;
  margin-left: 12px;
  border-radius: 8px;
  cursor: pointer;
  border: 1px solid #d7e2f2;
}

.login-row {
  display: flex;
  justify-content: space-between;
  align-items: center;
  margin: 2px 0 18px;
}

.register-link {
  color: #4a90e2;
  font-size: 13px;
}

.submit-btn {
  width: 100%;
}
</style>
