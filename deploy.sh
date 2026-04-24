#!/bin/bash

# Configuration from host info.md
REMOTE_USER="root"
REMOTE_HOST="101.133.160.196"
REMOTE_PASS="Lipangbo121"
REMOTE_DIR="/www/wwwroot/www.officeshichang.icu"

# Ensure deploy directory exists
mkdir -p deploy

# Package files
echo "Packaging project files..."
tar -czf project.tar.gz \
    --exclude='.venv' \
    --exclude='__pycache__' \
    --exclude='.git' \
    --exclude='.DS_Store' \
    --exclude='project.tar.gz' \
    --exclude='deploy.sh' \
    .

# Upload using expect
echo "Uploading project to $REMOTE_HOST..."
/usr/bin/expect <<EOF
set timeout 300
spawn scp project.tar.gz $REMOTE_USER@$REMOTE_HOST:/tmp/
expect {
    "password:" { send "$REMOTE_PASS\r"; exp_continue }
    "(yes/no)?" { send "yes\r"; exp_continue }
    eof
}
spawn ssh $REMOTE_USER@$REMOTE_HOST "mkdir -p /www/server/panel/vhost/cert/officeshichang.icu/"
expect {
    "password:" { send "$REMOTE_PASS\r"; exp_continue }
    "(yes/no)?" { send "yes\r"; exp_continue }
    eof
}
spawn scp certs_temp/fullchain.pem $REMOTE_USER@$REMOTE_HOST:/www/server/panel/vhost/cert/officeshichang.icu/
expect {
    "password:" { send "$REMOTE_PASS\r"; exp_continue }
    "(yes/no)?" { send "yes\r"; exp_continue }
    eof
}
spawn scp certs_temp/privkey.key $REMOTE_USER@$REMOTE_HOST:/www/server/panel/vhost/cert/officeshichang.icu/
expect {
    "password:" { send "$REMOTE_PASS\r"; exp_continue }
    "(yes/no)?" { send "yes\r"; exp_continue }
    eof
}
EOF

# Execute remote commands
echo "Executing remote deployment commands..."
/usr/bin/expect <<EOF
set timeout 300
spawn ssh $REMOTE_USER@$REMOTE_HOST "mkdir -p $REMOTE_DIR && tar -xzf /tmp/project.tar.gz -C $REMOTE_DIR && cd $REMOTE_DIR && \
    (apt-get update && apt-get install -y python3-venv || yum install -y python3-venv) && \
    python3 -m venv .venv && \
    ./.venv/bin/pip install -r requirements.txt && \
    cp deploy/time-test1.service /etc/systemd/system/ && \
    systemctl daemon-reload && \
    systemctl enable time-test1 && \
    systemctl restart time-test1 && \
    cp deploy/nginx.conf /www/server/panel/vhost/nginx/officeshichang.icu.conf && \
    /www/server/nginx/sbin/nginx -t && \
    /www/server/nginx/sbin/nginx -s reload && \
    find $REMOTE_DIR -not -name '.user.ini' -exec chown www:www {} + || find $REMOTE_DIR -not -name '.user.ini' -exec chown root:root {} + && \
    chmod -R 755 $REMOTE_DIR"
expect {
    "password:" { send "$REMOTE_PASS\r"; exp_continue }
    "(yes/no)?" { send "yes\r"; exp_continue }
    eof
}
EOF

# Cleanup
rm project.tar.gz
echo "Deployment finished!"
