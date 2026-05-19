import React, { useState } from 'react';
import { Layout, Dropdown, Avatar, Space, Typography, Tooltip, Button } from 'antd';
import {
  UserOutlined, LogoutOutlined, PlaySquareOutlined,
  GlobalOutlined, EyeOutlined, VideoCameraOutlined,
} from '@ant-design/icons';
import { useNavigate, useLocation } from 'react-router-dom';
import { ShowAndTellPanel } from '../features/showAndTell';
import { useAuth } from '../context/AuthContext';
import ScreenRecorder from '../components/ScreenRecorder';
import type { MenuProps } from 'antd';

const { Header, Content } = Layout;
const { Text } = Typography;

const REDWOOD = '#C74634';

const MainLayout: React.FC<{ children: React.ReactNode }> = ({ children }) => {
  const { user, logout } = useAuth();
  const navigate         = useNavigate();
  const location         = useLocation();
  const [showAndTellOpen, setShowAndTellOpen] = useState(false);

  const userMenu: MenuProps = {
    items: [
      { key: 'name',   label: <Text strong>{user?.name || user?.username}</Text>, disabled: true },
      { type: 'divider' },
      { key: 'logout', label: 'Sign Out', icon: <LogoutOutlined />, danger: true },
    ],
    onClick: ({ key }) => { if (key === 'logout') logout(); },
  };

  const navBtn = (path: string, icon: React.ReactNode, label: string) => {
    const active = location.pathname.startsWith(path);
    return (
      <Tooltip title={label} placement="bottom">
        <Button
          type={active ? 'primary' : 'text'}
          icon={icon}
          style={{
            color: active ? '#fff' : REDWOOD,
            background: active ? REDWOOD : 'transparent',
            borderColor: active ? REDWOOD : 'transparent',
            fontWeight: active ? 600 : 400,
          }}
          onClick={() => navigate(path)}
        >
          {label}
        </Button>
      </Tooltip>
    );
  };

  return (
    <Layout style={{ minHeight: '100vh' }}>
      <Header style={{
        background: '#fff',
        borderBottom: `2px solid ${REDWOOD}`,
        display: 'flex',
        alignItems: 'center',
        padding: '0 16px',
        gap: 8,
        position: 'sticky',
        top: 0,
        zIndex: 100,
      }}>
        {/* Logo */}
        <div style={{ marginRight: 24, display: 'flex', alignItems: 'center', gap: 8 }}>
          <div style={{ width: 28, height: 28, background: REDWOOD, borderRadius: 6, display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
            <VideoCameraOutlined style={{ color: '#fff', fontSize: 16 }} />
          </div>
          <Text strong style={{ fontSize: 16, color: REDWOOD, letterSpacing: '-0.3px' }}>ERP Studio</Text>
        </div>

        {/* Navigation */}
        <Space size={4}>
          {navBtn('/oracle-fusion', <GlobalOutlined />,    'Oracle Fusion')}
          {navBtn('/training',      <PlaySquareOutlined />, 'Training Library')}
        </Space>

        {/* Right toolbar */}
        <div style={{ marginLeft: 'auto', display: 'flex', alignItems: 'center', gap: 8 }}>
          <ScreenRecorder />

          <Tooltip title="Show & Tell — guided tours" placement="bottom">
            <Button
              type="text"
              icon={<EyeOutlined />}
              style={{ color: REDWOOD }}
              onClick={() => setShowAndTellOpen(true)}
            >
              Show &amp; Tell
            </Button>
          </Tooltip>

          <Dropdown menu={userMenu} placement="bottomRight" arrow>
            <div style={{ cursor: 'pointer', display: 'flex', alignItems: 'center', gap: 8, padding: '4px 8px', borderRadius: 6, border: '1px solid #e5e5e5' }}>
              {user?.photo
                ? <Avatar size={28} src={user.photo} />
                : <Avatar size={28} style={{ background: REDWOOD }} icon={<UserOutlined />} />}
              <Text style={{ fontSize: 13 }}>{user?.name || user?.username || 'User'}</Text>
            </div>
          </Dropdown>
        </div>
      </Header>

      <Content style={{ background: '#f5f5f5' }}>
        {children}
      </Content>

      <ShowAndTellPanel open={showAndTellOpen} onClose={() => setShowAndTellOpen(false)} />
    </Layout>
  );
};

export default MainLayout;
