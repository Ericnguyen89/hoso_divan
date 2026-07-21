# Front-end — mô hình deploy kiểu production (GitOps + ArgoCD + ghcr)

Nâng cấp từ "docker pull + kubectl apply thủ công" sang pipeline production thực tế,
**phù hợp cluster VPN-only** (API k8s không ra Internet).

## 1. Nguyên tắc production
1. **Immutable image tag**: tag theo commit `:<git-sha>` (hoặc semver `:v1.2.3`), KHÔNG dùng
   `:dev`/`:latest` mutable. Biết chính xác cái gì đang chạy + rollback = trỏ lại tag cũ.
2. **GitOps = git là nguồn sự thật**: trạng thái mong muốn (manifest + tag) nằm trong git;
   ArgoCD trong cluster tự đồng bộ. Mọi thay đổi qua git → audit + rollback bằng `git revert`.
3. **Pull-based CD** (bắt buộc với cluster VPN-only): ArgoCD/Image-Updater CHẠY TRONG cluster,
   tự KÉO từ git/ghcr. CI cloud chỉ build+push, KHÔNG cần chạm tới API k8s.
4. **Tách config theo môi trường**: base + overlays (kustomize); secret KHÔNG để git thô
   (dùng sealed-secrets/external-secrets, hoặc tạo out-of-band).
5. **Promote cùng 1 image** dev→staging→prod (đừng build lại mỗi env) — TRỪ biến client bake
   lúc build (`NEXT_PUBLIC_*`): nếu khác theo env thì phải build riêng, hoặc chuyển sang
   **runtime config** (đọc lúc chạy) để 1 image chạy mọi env.

## 2. Kiến trúc cho cluster này
```
[FE repo GitHub]
    │ push nhánh/tag
    ▼
[GitHub Actions]  (CI - cloud)
    │ build+test → docker build → push ghcr.io/org/fe:<sha> (+ moving tag)
    │ (write-back) cập nhật tag ảnh vào [GitOps repo]  (git commit/PR)
    ▼
[GitOps repo] (manifest + image tag = nguồn sự thật)
    ▲ pull
[ArgoCD trong cluster] ── sync ──► Deployment/Service/Ingress + cert
    ├─ fe-dev      → app.sd-dev.mikotechdev.net
    ├─ fe-staging  → staging.tastelife.app
    └─ fe-prod     → tastelife.app (khi có)
```
CI cloud **không** rollout; ArgoCD (in-cluster) kéo về → chạy tốt dù API k8s chỉ trong VPN.

## 3. Bố cục repo (kustomize base + overlays)
```
gitops/apps/frontend/
  base/               kustomization + deployment/service/ingress (dùng chung)
  overlays/dev/       patch: host app.sd-dev..., image tag, env dev, replicas
  overlays/staging/   patch: host staging.tastelife.app, env staging
  overlays/prod/      patch: host tastelife.app, replicas 2+, env prod
argocd/
  fe-dev.yaml  fe-staging.yaml  fe-prod.yaml   (ArgoCD Application, mỗi env 1 cái)
```
Mỗi ArgoCD Application trỏ 1 overlay + auto-sync {prune, selfHeal}.

## 4. CI (GitHub Actions) — build + tag bất biến + write-back
`.github/workflows/deploy.yml` (rút gọn):
```yaml
on: { push: { branches: [develop, staging], tags: ['v*'] } }
jobs:
  build:
    runs-on: ubuntu-latest
    permissions: { contents: write, packages: write }
    steps:
      - uses: actions/checkout@v4
      - id: env
        run: |   # nhánh -> env + API URL client (bake)
          case "${{ github.ref_name }}" in
            staging) echo "env=staging; api=https://gw.sd-staging.mikotechdev.net" ;;
            v*)      echo "env=prod;    api=https://gw.tastelife.app" ;;
            *)       echo "env=dev;     api=https://gw.sd-dev.mikotechdev.net" ;;
          esac >> $GITHUB_OUTPUT
      - uses: docker/login-action@v3
        with: { registry: ghcr.io, username: ${{ github.actor }}, password: ${{ secrets.GITHUB_TOKEN }} }
      - uses: docker/build-push-action@v6
        with:
          context: .
          push: true
          build-args: NEXT_PUBLIC_API_URL=${{ steps.env.outputs.api }}
          tags: ghcr.io/${{ github.repository_owner }}/sd-web:${{ github.sha }}
      # WRITE-BACK: cập nhật tag vào GitOps repo -> ArgoCD tự sync
      - run: |
          git clone https://x:${{ secrets.GITOPS_PAT }}@github.com/OWNER/Deploy-VPS gp && cd gp
          cd gitops/apps/frontend/overlays/${{ steps.env.outputs.env }}
          kustomize edit set image ghcr.io/OWNER/sd-web=ghcr.io/OWNER/sd-web:${{ github.sha }}
          git commit -am "fe/${{ steps.env.outputs.env }}: ${{ github.sha }}" && git push
```

## 5. CD — ArgoCD Application (pull-based)
`argocd/fe-staging.yaml`:
```yaml
apiVersion: argoproj.io/v1alpha1
kind: Application
metadata: { name: fe-staging, namespace: argocd }
spec:
  project: platform
  source:
    repoURL: https://github.com/OWNER/Deploy-VPS
    targetRevision: main
    path: gitops/apps/frontend/overlays/staging
  destination: { server: https://kubernetes.default.svc, namespace: sd-staging-web }
  syncPolicy:
    automated: { prune: true, selfHeal: true }
    syncOptions: [ CreateNamespace=true ]
```
Sau khi CI write-back tag mới → ArgoCD phát hiện diff → sync → rollout. **Rollback**: `git revert`
commit tag trong GitOps repo.

**Thay cho write-back — ArgoCD Image Updater** (không sửa git từ CI):
cài image-updater trong cluster + annotation trên Application để nó theo dõi ghcr và tự cập nhật:
```yaml
metadata:
  annotations:
    argocd-image-updater.argoproj.io/image-list: web=ghcr.io/OWNER/sd-web
    argocd-image-updater.argoproj.io/web.update-strategy: newest-build
```

## 6. Secret trong production (đừng để git thô)
- **sealed-secrets** (Bitnami): mã hoá secret → commit git an toàn → controller giải mã trong cluster. Hợp GitOps nhất.
- **external-secrets**: kéo từ Vault/cloud secret manager lúc chạy.
- Tối thiểu (như hiện tại): tạo Secret out-of-band bằng `kubectl` (STRAPI_API_TOKEN…), git chỉ tham chiếu tên.

## 7. Quy trình vận hành
| Hành động | Cách làm |
|---|---|
| Deploy dev | push nhánh `develop` → CI build → ArgoCD tự deploy `app.sd-dev` |
| Lên staging | merge/push `staging` → tự deploy `staging.tastelife.app` |
| Lên prod | tạo tag `vX.Y.Z` → build + write-back overlay prod → ArgoCD deploy |
| Rollback | `git revert` commit tag trong GitOps repo (ArgoCD tự đưa về bản cũ) |
| Xem trạng thái | ArgoCD UI (argocd.sd-dev.mikotechdev.net) hoặc `kubectl -n argocd get app` |

## 8. Migration từ hiện tại (thủ công) → GitOps
1. Chuyển manifest FE (`deploy/frontend/k8s/*`) thành `gitops/apps/frontend/{base,overlays}` (kustomize).
2. Thêm CI immutable-tag + (write-back hoặc Image Updater).
3. Tạo 3 ArgoCD Application (dev/staging/prod).
4. `kubectl delete` các Deployment/Ingress tạo tay (ArgoCD sẽ tạo lại từ git — hoặc để adopt).
5. Từ đó chỉ thao tác qua git; không `kubectl apply` tay nữa.
```
